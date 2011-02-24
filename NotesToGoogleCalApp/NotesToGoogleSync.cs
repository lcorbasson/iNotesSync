using System;
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
        [System.Diagnostics.DebuggerNonUserCodeAttribute()] 
        public int Sync(BackgroundWorker worker, DoWorkEventArgs e)
        {
            int retVal = 0;
            StreamReader xmlStream;
            List<EventEntry> newEvents = new List<EventEntry>();
            int listCount, intval, progressCount;

            if (scGoogleService==null)
            {   // Google service is Null, cannot continue
                if (worker != null) { throw new Exception("Google Service not initialized.\r\n" + scGoogleService.LastError); }
                return SYNCNULLGSERVICE;
            }
            else if (scNotesService == null)
            {   // Notes service is Null, cannot continue
                if (worker != null) { throw new Exception("Notes Service not initialized.\r\n" + scNotesService.LastError); }
                return SYNCNULLNSERVICE;
            }
            else // Services should be good, do work.
            {
                // Remove old notes entries from the calendar
                if (scGoogleService.ClearEventEntriesBySearch("Created from [LNotes]") < 0)
                {   // There was an error clearing calendar
                    if (worker != null) { throw new Exception("Error deleting current Notes Entries.\r\n" + scGoogleService.LastError); }
                    return SYNCFAILCLEAR;
                }

                if (worker != null) { worker.ReportProgress(20); }          // Update progress bar; if backgorund process exists

                // Get the XML stream
                xmlStream = scNotesService.GetNotesXMLCalendar();
                if (xmlStream == null) { throw new Exception("Unable to get XML stream from Domino; check URL and Login information."); }
                if (worker != null) { worker.ReportProgress(25); }          // Update progress bar; if background process exists

                // Parse the XML into a List of EventEntries
                newEvents = ConvertXMLtoEvents(xmlStream);
                listCount = newEvents.Count;
                if (worker != null) { worker.ReportProgress(30); }          // Update progress bar; if background process exists

                intval = 0;
                foreach (EventEntry ent in newEvents)
                {
                    intval++;

                    if (scGoogleService.InsertEvent(ent) != null)
                    {
                        retVal++;
                        progressCount = ((70 / listCount) * intval) + 30;
                        if (worker != null) { worker.ReportProgress(progressCount); }
                    }
                    else
                    {
                        if (worker != null) { throw new Exception("Unable to insert event. " + scGoogleService.LastError); }
                        return SYNCFAILURE;
                    }
                }
            }

            xmlStream.Close();
            scNotesService.CloseConnection();

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
        [System.Diagnostics.DebuggerNonUserCodeAttribute()] 
        private List<EventEntry> ConvertXMLtoEvents(StreamReader _stream)
        {
            //MessageBox.Show("Created XML reader", "xml", MessageBoxButtons.OK);
            List<EventEntry> events = new List<EventEntry>();   // List of events to return
            XmlTextReader xmlReader = new XmlTextReader(_stream);   // XMLTextReader from passed stream
            xmlReader.WhitespaceHandling = WhitespaceHandling.All;
            //XmlReader xmlReader = XmlReader.Create(_stream);
            Boolean fulldayevent, remindevent;
            //MessageBox.Show("Done creating reader", "xml", MessageBoxButtons.OK);

            EventEntry newEvent = null;
            String sValue;                              // Short term holding of datetime string
            DateTime startDateTime = DateTime.Now;      // Long term holding of datetime string
            DateTime fullDateTime = DateTime.Now;
            DateTime endDateTime = startDateTime.AddHours(1);
            When eventWhen;
        
            // Read Elements (forced top level)
            xmlReader.ReadToFollowing("viewentry");
            //xmlReader.Read();
            while (!xmlReader.EOF)
            {
                //MessageBox.Show("Reading <viewentry>", "xml", MessageBoxButtons.OK);
                fulldayevent = false;
                remindevent = false;

                // <viewentry>
                if (xmlReader.IsStartElement() && xmlReader.Name == "viewentry")
                {
                    newEvent = new EventEntry();    // Create new event for this viewentry                                                                   
                    newEvent.Title.Text = "";

                    // <entrydata>
                    for (int i = 0; i < 6; i++)
                    {                        
                        xmlReader.ReadToFollowing("entrydata");
                        //MessageBox.Show("Reading <entrydata>", "xml", MessageBoxButtons.OK);

                        if (xmlReader.IsStartElement() && xmlReader.Name == "entrydata")
                        {   
                            // Read the Name attribute of <entrydata>
                            xmlReader.MoveToAttribute("name");
                            //MessageBox.Show("Reading attribute name", "xml", MessageBoxButtons.OK);

                            // Switch based on what name=??
                            switch (xmlReader.Value)
                            {
                                // 134 - start date from datetime.value
                                case FULLDAYEVENTFIELD:
                                    xmlReader.ReadToFollowing("datetime");
                                    xmlReader.Read();

                                    sValue = xmlReader.Value;
                                    //MessageBox.Show("fullday " + sValue + " is " + sValue.Length.ToString(), "sValue", MessageBoxButtons.OK);

                                    // Create the FullDateTime object used to set date
                                    fullDateTime = new DateTime(Convert.ToInt32(sValue.Substring(0, 4)),
                                        Convert.ToInt32(sValue.Substring(4, 2)), Convert.ToInt32(sValue.Substring(6, 2)),
                                        Convert.ToInt32(sValue.Substring(9, 2)), Convert.ToInt32(sValue.Substring(11, 2)),
                                        Convert.ToInt32(sValue.Substring(13, 2)), DateTimeKind.Utc);

                                    // We need to manimpulate the time using the last 3 digits of the time code sent making it the correct UTC time
                                    fullDateTime = OffSetUTCTime(fullDateTime, sValue.Substring(18,3));

                                    break;
                                // 144 - start date from datetime.value
                                case STARTDATETIMEFIELD:
                                    if (!fulldayevent)
                                    {
                                        xmlReader.ReadToFollowing("datetime");
                                        xmlReader.Read();

                                        sValue = xmlReader.Value;
                                        //MessageBox.Show("Start " + sValue + " is " + sValue.Length.ToString(), "sValue", MessageBoxButtons.OK);

                                        startDateTime = new DateTime(Convert.ToInt32(sValue.Substring(0, 4)),
                                            Convert.ToInt32(sValue.Substring(4, 2)), Convert.ToInt32(sValue.Substring(6, 2)),
                                            Convert.ToInt32(sValue.Substring(9, 2)), Convert.ToInt32(sValue.Substring(11, 2)),
                                            Convert.ToInt32(sValue.Substring(13, 2)), DateTimeKind.Utc);

                                        // We need to manimpulate the time using the last 3 digits of the time code sent making it the correct UTC time
                                        startDateTime = OffSetUTCTime(startDateTime, sValue.Substring(18, 3));
                                    }
                                    else
                                    {
                                        startDateTime = fullDateTime;
                                    }
                                    break;
                                // 146 - end date from datetime.value
                                case ENDDATETIMEFIELD:
                                    if (!(fulldayevent || remindevent))
                                    {
                                        xmlReader.ReadToFollowing("datetime");
                                        xmlReader.Read();

                                        sValue = xmlReader.Value;
                                        //MessageBox.Show("End " + sValue + " is " + sValue.Length.ToString(), "sValue", MessageBoxButtons.OK);

                                        endDateTime = new DateTime(Convert.ToInt32(sValue.Substring(0, 4)),
                                            Convert.ToInt32(sValue.Substring(4, 2)), Convert.ToInt32(sValue.Substring(6, 2)),
                                            Convert.ToInt32(sValue.Substring(9, 2)), Convert.ToInt32(sValue.Substring(11, 2)),
                                            Convert.ToInt32(sValue.Substring(13, 2)), DateTimeKind.Utc);

                                        // We need to manimpulate the time using the last 3 digits of the time code sent making it the correct UTC time
                                        endDateTime = OffSetUTCTime(endDateTime, sValue.Substring(18, 3));
                                    }
                                    else
                                    {
                                        endDateTime = fullDateTime.AddHours(23);
                                    }
                                    break;
                                // 149 - 160-appoint 158-meeting 177-holiday 10-reminder
                                case ENTRYTYPEFIELD:
                                    xmlReader.ReadToFollowing("number");
                                    xmlReader.Read();

                                    //MessageBox.Show("type " + sValue + " is " + sValue.Length.ToString(), "sValue", MessageBoxButtons.OK);

                                    if ((xmlReader.Value == "160") || (xmlReader.Value == "205"))
                                    {   // This is an Appointment
                                        newEvent.Title.Text += "App: ";
                                    }
                                    else if ((xmlReader.Value == "158") || (xmlReader.Value=="210"))
                                    {   // This is a Meeting
                                        newEvent.Title.Text += "Meet: ";
                                    }
                                    else if (xmlReader.Value == "177")
                                    {
                                        newEvent.Title.Text += "Hol: ";
                                        fulldayevent = true;
                                    }
                                    else if ((xmlReader.Value == "194") || (xmlReader.Value == "199"))
                                    {
                                        newEvent.Title.Text += "Full: ";
                                        fulldayevent = true;
                                    }
                                    else if ((xmlReader.Value == "10") || (xmlReader.Value == "195"))
                                    {   // This is a full day reminder type
                                        newEvent.Title.Text += "Rem: ";
                                        remindevent = true;
                                    }
                                    break;
                                // 147 - go to text elements
                                case CONTENTFIELD:

                                    xmlReader.Read();   // <textlist>
                                    int textcount = 0;
                                    do
                                    {
                                        if (xmlReader.IsStartElement() && xmlReader.Name == "text")
                                        {
                                            xmlReader.Read();   // Text element within <text>
                                            //MessageBox.Show(xmlReader.Value, "value", MessageBoxButtons.OK);
                                            
                                            if (xmlReader.Value.ToString().Contains("Location:"))
                                            {   // Set the event Location to the Location string
                                                newEvent.Locations.Add(new Where("", "", xmlReader.Value));
                                            }
                                            else if (xmlReader.Value.ToString().Contains("Chair:"))
                                            {   // Add the Chair contents to Summary
                                                newEvent.Content.Content += xmlReader.Value+"\r\n";
                                            }
                                            else
                                            {   // Add the text value to summary
                                                switch (textcount)
                                                {
                                                    case 0:
                                                        newEvent.Title.Text += xmlReader.Value;
                                                        break;
                                                    case 1:
                                                        newEvent.Locations.Add(new Where("", "", xmlReader.Value));
                                                        break;
                                                    default:
                                                        newEvent.Content.Content += xmlReader.Value + "\r\n";
                                                        break;
                                                }                                                                                                
                                            }
                                            textcount++;
                                        }
                                        xmlReader.Read();
                                    } while ((xmlReader.Name != "textlist") && (xmlReader.Name != "entrydata")); // </textlist> or </entrydata>

                                    xmlReader.Read();
                                    break;
                                default:
                                    // Possibly add more choices later
                                    break;
                            } // end switch
                        } // end if
                    } // </entrydata>

                    // Insert the newEvent
                    newEvent.Content.Content += "\r\nCreated from [LNotes]";

                    if (fulldayevent)
                    {
                        eventWhen = new When();
                        eventWhen.StartTime = fullDateTime;
                        eventWhen.AllDay = true;
                    }
                    else
                    {
                        if (remindevent)
                        {
                            eventWhen = new When(startDateTime, startDateTime.AddHours(1));
                        }
                        else
                        {
                            eventWhen = new When(startDateTime, endDateTime);
                        }
                    }

                    newEvent.Times.Add(eventWhen);                    
                    events.Add(newEvent);
                    newEvent = null;
                }

                xmlReader.ReadToFollowing("viewentry");     // Go to the next <viewentry>
            } // </viewEntry>

            xmlReader.Close();
            return events;
        }

        /// <summary>
        /// This method is used to add or subtract the appropriate time values making the current time UTC correct
        /// </summary>
        /// <param name="current"></param>
        /// <param name="UTCValue"></param>
        /// <returns></returns>
        private DateTime OffSetUTCTime(DateTime current, String UTCValue)
        {
            DateTime updated = current;

            try
            {
                Int16 UTCHours = System.Convert.ToInt16(UTCValue.Substring(1, 2));
                //Int16 UTCHours = System.Convert.ToInt16(UTCValue);

                if (UTCValue.Contains('-'))
                {
                    updated = current.AddHours(UTCHours);
                }
                else
                {
                    updated = current.AddHours(-1 * UTCHours);
                }
            }
            catch (FormatException fe)
            {   

            }
            
            return updated;
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
        private const string FULLDAYEVENTFIELD = "$134";
    }
}
