﻿using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Threading;
using System.Xml;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;
using Google.GData.AccessControl;  

namespace NotesToGoogle
{
    /// <summary>
    /// NotesToGoogleSync class extends the Thread class.  It is run as an extra thread and is responsible
    /// for the actual syncing between Notes and Google.
    /// </summary>
    class NotesToGoogleSync
    {
        /// <summary>
        /// Constructor with no parameters
        /// </summary>
        public NotesToGoogleSync()
        {
            scGoogleService = null;
            scNotesService = null;
        }

        /// <summary>
        /// Constructor with GoogleServiceConnect and NotesServiceConnect parameters
        /// </summary>
        /// <param name="_googleService">Object representing the googleServiceConnector</param>
        /// <param name="_notesService">Object representing the notesServiceConnector</param>
        public NotesToGoogleSync(GoogleServiceConnect _googleService, NotesServiceConnect _notesService)
        {
            scGoogleService = _googleService;
            scNotesService = _notesService;
        }

        public int Sync()
        {
            return this.Sync((BackgroundWorker)null, (DoWorkEventArgs)null);
        }

        /// <summary>
        /// Method called to initiate the sync process.
        /// </summary>
        /// <returns>Returns the success value</returns>
        public int Sync(BackgroundWorker worker, DoWorkEventArgs e)
        {
            int retVal = 0;
            Stream xmlStream;
            List<EventEntry> newEvents = new List<EventEntry>();
            int listCount, intval, progressCount;

            if (scGoogleService==null)
            {   // Google service is Null, cannot continue
                if (worker != null) { throw new Exception("Google Service not initialized."); }
                return SYNCNULLGSERVICE;
            }
            else if (scNotesService == null)
            {   // Notes service is Null, cannot continue
                if (worker != null) { throw new Exception("Notes Service not initialized."); }
                return SYNCNULLNSERVICE;
            }
            else // Services should be good, do work.
            {
                // Remove old notes entries from the calendar
                if (scGoogleService.ClearEventEntriesBySearch("Created from [LNotes]") <= 0)
                {   // There was an error clearing calendar
                    if (worker != null) { throw new Exception("Error deleting current Notes Entries."); }
                    return SYNCFAILCLEAR;
                }

                if (worker != null) { worker.ReportProgress(20); }

                // Get the XML stream
                xmlStream = scNotesService.GetNotesXMLCalendar();

                if (worker != null) { worker.ReportProgress(25); }

                // Parse the XML into a List of EventEntries
                newEvents = ConvertXMLtoEvents(xmlStream);
                listCount = newEvents.Count;

                intval = 0;
                foreach (EventEntry ent in newEvents)
                {
                    intval++;

                    if (scGoogleService.InsertEvent(ent) != null)
                    {
                        retVal++;
                        progressCount = ((75 / listCount) * intval) + 25;
                        if (worker != null) { worker.ReportProgress(progressCount); }
                    }
                    else
                    {
                        if (worker != null) { throw new Exception("Unable to insert event."); }
                        return SYNCFAILURE;
                    }
                }
            }

            return retVal;
        }

        /// <summary>
        /// Accessor method for GoogleService
        /// </summary>
        public GoogleServiceConnect GoogleService
        {
            get
            {
                return scGoogleService;
            }
            set
            {
                scGoogleService = value;
            }
        }

        /// <summary>
        /// Accessor method for NotesService
        /// </summary>
        public NotesServiceConnect NotesService
        {
            get
            {
                return scNotesService;
            }
            set
            {
                scNotesService = value;
            }
        }

        /// <summary>
        /// Method returns a list of EventEntries to be input into GoogleService
        /// </summary>
        /// <param name="_stream">XML stream passed in to create XMLTextReader</param>
        /// <returns>List of EventEntry objects</returns>
        private List<EventEntry> ConvertXMLtoEvents(Stream _stream)
        {
            List<EventEntry> events = new List<EventEntry>();   // List of events to return
            XmlTextReader xmlReader = new XmlTextReader(_stream);   // XMLTextReader from passed stream
            xmlReader.WhitespaceHandling = WhitespaceHandling.None;

            EventEntry newEvent = null;
            String sValue;                              // Short term holding of datetime string
            DateTime startDateTime = DateTime.Now;      // Long term holding of datetime string
            DateTime endDateTime = startDateTime.AddHours(1);

            // Read Elements (forced top level)
            xmlReader.ReadToFollowing("viewentry");
            while (!xmlReader.EOF)
            {
                // <viewentry>
                if (xmlReader.IsStartElement() && xmlReader.Name == "viewentry")
                {
                    newEvent = new EventEntry();    // Create new event for this viewentry                                               

                    // <entrydata>
                    for (int i = 0; i < 6; i++)
                    {
                        xmlReader.ReadToFollowing("entrydata");

                        if (xmlReader.IsStartElement() && xmlReader.Name == "entrydata")
                        {   // Read the Name attribute of <entrydata>
                            xmlReader.MoveToAttribute("name");

                            // Switch based on what name=??
                            switch (xmlReader.Value)
                            {
                                // 144 - start date from datetime.value
                                case STARTDATETIMEFIELD:
                                    xmlReader.ReadToFollowing("datetime");
                                    xmlReader.Read();

                                    sValue = xmlReader.Value;
                                    startDateTime = new DateTime(Convert.ToInt32(sValue.Substring(0, 4)),
                                        Convert.ToInt32(sValue.Substring(4, 2)), Convert.ToInt32(sValue.Substring(6, 2)),
                                        Convert.ToInt32(sValue.Substring(9, 2)), Convert.ToInt32(sValue.Substring(11, 2)),
                                        Convert.ToInt32(sValue.Substring(13, 2)), DateTimeKind.Local);
                                    break;
                                // 146 - end date from datetime.value
                                case ENDDATETIMEFIELD:
                                    xmlReader.ReadToFollowing("datetime");
                                    xmlReader.Read();

                                    sValue = xmlReader.Value;
                                    endDateTime = new DateTime(Convert.ToInt32(sValue.Substring(0, 4)),
                                        Convert.ToInt32(sValue.Substring(4, 2)), Convert.ToInt32(sValue.Substring(6, 2)),
                                        Convert.ToInt32(sValue.Substring(9, 2)), Convert.ToInt32(sValue.Substring(11, 2)),
                                        Convert.ToInt32(sValue.Substring(13, 2)), DateTimeKind.Local);
                                    break;
                                // 149 - 160-appoint 158-meeting
                                case ENTRYTYPEFIELD:
                                    xmlReader.ReadToFollowing("number");
                                    xmlReader.Read();

                                    if (xmlReader.Value == "160")
                                    {   // This is an Appointment
                                        newEvent.Title.Text += "Appoint: ";
                                    }
                                    else if (xmlReader.Value == "158")
                                    {   // This is a Meeting
                                        newEvent.Title.Text += "Meeting: ";
                                    }
                                    break;
                                // 147 - go to text elements
                                case CONTENTFIELD:
                                    xmlReader.Read();   // <textlist>
                                    do
                                    {
                                        if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "text"))
                                        {
                                            xmlReader.Read();   // Text element within <text>
                                            if (xmlReader.Value.ToString().Contains("Location:"))
                                            {   // Set the event Location to the Location string
                                                newEvent.Locations.Add(new Where("", "", xmlReader.Value));
                                            }
                                            else if (xmlReader.Value.ToString().Contains("Chair:"))
                                            {   // Add the Chair contents to Summary
                                                newEvent.Summary.Text += xmlReader.Value;
                                            }
                                            else
                                            {   // Add the text value to summary
                                                newEvent.Title.Text += xmlReader.Value;
                                                newEvent.Content.Content += xmlReader.Value;
                                            }
                                        }
                                        xmlReader.Read();
                                    } while (xmlReader.Name != "textlist"); // </textlist>

                                    xmlReader.Read();
                                    break;
                                default:
                                    // Possibly add more choices later
                                    break;
                            }
                        }
                    } // </entrydata>

                    // Insert the newEvent
                    newEvent.Content.Content += "\r\n Created from [LNotes]";
                    newEvent.Times.Add(new When(startDateTime, endDateTime));
                    events.Add(newEvent);
                    newEvent = null;
                }

                xmlReader.ReadToFollowing("viewentry");     // Go to the next <viewentry>

            } // </viewEntry>

            return events;
        }

        // Private Class Variables
        private GoogleServiceConnect scGoogleService;
        private NotesServiceConnect scNotesService;

        // Private Constants :: Syncing status
        public const int SYNCSUCCESS = 0;
        public const int SYNCFAILURE = -1;
        public const int SYNCNULLGSERVICE = -2;
        public const int SYNCNULLNSERVICE = -3;
        public const int SYNCFAILCLEAR = -4;

        // Private Constants :: Searching through Notes
        private const string STARTDATETIMEFIELD = "$144";
        private const string ENDDATETIMEFIELD = "$146";
        private const string CONTENTFIELD = "$147";
        private const string ENTRYTYPEFIELD = "$149";
    }
}
