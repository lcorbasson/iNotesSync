﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;
using Google.GData.AccessControl;
using System.Windows.Forms;
using System.Net;   

namespace NotesToGoogle
{
    /// <summary>
    /// The GoogleServiceConnect class is used to make connections to google calendar; it can be used to:
    /// view, add, delete, update calendars and calendar entries
    /// </summary>
    class GoogleServiceConnect
    {
        /// <summary>
        /// Creates GoogleServiceConnect object, chained to constructor.
        /// </summary>
        /// <param name="_GoogleLogin">Supplies Google Login / Default calender</param>
        /// <param name="_GooglePassword">Supplies Google Password</param>
        public GoogleServiceConnect(String _GoogleLogin, String _GooglePassword) :
            this(_GoogleLogin, _GooglePassword, true, false) { }

        /// <summary>
        /// Creates GoogleServiceConnect object, chained to constructor.
        /// </summary>
        /// <param name="_GoogleLogin">Supplies Google Login / Default calendar</param>
        /// <param name="_GooglePassword">Supplies Google password</param>
        /// <param name="_UseSSL">Supplies Use SSL Switch</param>
        public GoogleServiceConnect(String _GoogleLogin, String _GooglePassword, Boolean _UseSSL) :
            this(_GoogleLogin, _GooglePassword, _UseSSL, false) { }

        /// <summary>
        /// Creates GoogleServiceConnect object.
        /// </summary>
        /// <param name="_GoogleLogin">Supplies Google Login / Default calendar</param>
        /// <param name="_GooglePassword">Supplies Google Password</param>
        /// <param name="_UseSSL">Supplies Use SSL Switch</param>
        /// <param name="_notifications">Supplies notification switch</param>
        public GoogleServiceConnect(String _GoogleLogin, String _GooglePassword, Boolean _UseSSL, Boolean _notifications)
        {
            try
            {
                // Gather initial variables
                sGoogleLogin = _GoogleLogin;
                sGooglePassword = _GooglePassword;
                defaultCalendar = _GoogleLogin;
                bUseSSL = _UseSSL;
                bNotifications = _notifications;
                iDaysAhead = 14;

                csService = new CalendarService("NotesToGoogleApp");
                
                // Depending on the 
                if (bUseSSL)
                {
                    sHttpProtocol = "https://";
                }
                else
                {
                    sHttpProtocol = "http://";
                }

                // Create the Calendar Service
                

                // Set the Google Calendar Urls               
                sOwnerCalURI = sHttpProtocol + "www.google.com/calendar/feeds/default/owncalendars/full";

                // Set user credentials if they exist
                if (sGoogleLogin != null && sGoogleLogin.Length > 0)
                {
                    csService.setUserCredentials(sGoogleLogin, sGooglePassword);                    
                }
            }
            catch (ServiceUnavailableException e)
            {
                return;
            }
        }

        /// <summary>
        /// CreateCalendar is used to create a new calendar to the owners profile.
        /// </summary>
        /// <param name="_calendarName">Name of the new calendar</param>
        /// <returns>CalendarEntry object of the new calendar</returns>
        public CalendarEntry CreateCalendar(String _calendarName)
        {
            CalendarEntry createdCalendar;
            Uri postUri;

            try
            {
                CalendarEntry calendar = new CalendarEntry();
                calendar.Title.Text = _calendarName;
                calendar.Summary.Text = "Calendar sync from Lotus Notes.";
                calendar.TimeZone = "America/Toronto";
                // Currently this is returning an error
                // calendar.Hidden = false;
                calendar.Color = "#2952A3";
                calendar.Location = new Where("", "", "Toronto");

                postUri = new Uri(sOwnerCalURI);                
                createdCalendar = (CalendarEntry)csService.Insert(postUri, calendar);                
            }
            catch (GDataRequestException ex)
            {
                return (CalendarEntry) null;
            }
            return createdCalendar;
        }

        /// <summary>
        /// Deletes a calendar with the specified name.
        /// </summary>
        /// <param name="_calendarName">Calendar name to delete</param>
        /// <returns>Boolean whether calendar was deleted</returns>
        public System.Boolean DeleteCalendar(String _calendarName)
        {
            Boolean isDeleted = false;
            CalendarEntry cal;

            try
            {
                cal = GetCalendarByName(_calendarName);             
                cal.Delete();
                isDeleted = true;                                
            }
            catch (Exception e)
            {
                return isDeleted;
            }

            return isDeleted;
        }

        /// <summary>
        /// Method used to clear out all Lotus Notes specific entries; Lotus notes specific entries are 
        /// denoted by a NL: in the subject field.  Calendar entries will only be deleted if they contain
        /// the string 'Created from [LNotes]'; so when creating new entries FROM Lotus ensure that this
        /// string is somewhere within the data.
        /// </summary>
        /// <param name="_calendar">Calendar name to search for</param>
        /// <param name="_searchText">String to search entries for to delete</param>
        /// <returns>Number of entries deleted</returns>
        public System.Int32 ClearEventEntriesBySearch(String _calendar, String _searchText)
        {
            CalendarEntry calToSearch;
            EventQuery query;
            EventFeed calFeed, batchFeed;
            DateTime today = DateTime.Now;

            try
            {
                // Create the query
                calToSearch = GetCalendarByName(_calendar);
                                
                query = new EventQuery(GetAlternateURL(calToSearch));
                query.Query = _searchText;
                query.StartDate = today;
                query.EndDate = today.AddDays(iDaysAhead);
                
                calFeed = csService.Query(query);
                foreach (EventEntry ev in calFeed.Entries)
                {
                    ev.BatchData = new GDataBatchEntryData("D", GDataBatchOperationType.delete); 
                }

                batchFeed = (EventFeed)csService.Batch(calFeed, new Uri(calFeed.Batch));
            }
            catch (Exception ex)
            {                
                return -1;
            }          
            return calFeed.Entries.Count;
        }

        /// <summary>
        /// Method used to clear out all Lotus Notes specific entries; Lotus notes specific entries are 
        /// denoted by a NL: in the subject field.  Calendar entries will only be deleted if they contain
        /// the string 'Created from [LNotes]'; so when creating new entries FROM Lotus ensure that this
        /// string is somewhere within the data.
        /// </summary>
        /// <param name="_searchText">Search string to choose entries to delete</param>
        /// <returns>Number of events deleted</returns>
        public int ClearEventEntriesBySearch(String _searchText)
        {
            return ClearEventEntriesBySearch(defaultCalendar, _searchText);
        }

        /// <summary>
        /// Method used to add a number of events to 
        /// </summary>
        /// <param name="_calendar">Calendar to insert entries into</param>
        /// <param name="events">List of evententries to add</param>
        /// <returns>Whether or not the batch add worked properly</returns>
        public Boolean BatchAddEvents(CalendarEntry _calendar, List<EventEntry> events)
        {


            return true;
        }

        /// <summary>
        /// Returns the calendar entry object by the specified name
        /// </summary>
        /// <param name="_calendarName">Name of calendar to return</param>
        /// <returns>Calendar with specified name</returns>
        public CalendarEntry GetCalendarByName(String _calendarName)
        {
            CalendarEntry returnCalendar = null;
            CalendarFeed calendars;
            CalendarQuery query = new CalendarQuery();

            try
            {
                // Create the query to the Ownders (all calendars) URL
                query.Uri = new Uri(sOwnerCalURI);
                calendars = (CalendarFeed)csService.Query(query);

                // Seach through all Calendars in the CalendarFeed
                foreach (CalendarEntry cal in calendars.Entries)
                {
                    // If the calendar title matches the name to delete; delete it!  
                    if (cal.Title.Text.Equals(_calendarName))
                    {
                        returnCalendar = cal;
                    }
                }
            }
            catch (Exception e)
            {
                return (CalendarEntry)null;
            }

            return returnCalendar;
        }

        /// <summary>
        /// Test the connection to the GoogleService
        /// </summary>
        /// <returns>returns True for success, false for failure</returns>
        public Boolean TestConnection()
        {
            try
            {
                CalendarFeed calendars;
                CalendarQuery query = new CalendarQuery();

                query.Uri = new Uri(sOwnerCalURI);
                calendars = (CalendarFeed)csService.Query(query);
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Get / Set method for the private dtStartDate variable
        /// GoogleServiceConnect.StartDate = value to set
        /// </summary>
        public DateTime StartDate
        {
            get
            {
                return dtStartDate;
            }
            set
            {
                dtStartDate = value;
            }
        }

        /// <summary>
        /// Get / Set accessor method for the dtEndDate private variable
        /// GoogleServiceConnect.EndDate = value to set
        /// </summary>
        public DateTime EndDate
        {
            get
            {
                return dtEndDate;
            }
            set
            {
                dtEndDate = value;
            }
        }

        /// <summary>
        /// Method used to return the Alternate Link string value.  This URL is required when doing
        /// any type of searching.
        /// </summary>
        /// <param name="_calToSearch">CalendarEntry to return alternate link for.</param>
        /// <returns>Alternate link URL as string.</returns>
        private String GetAlternateURL(CalendarEntry _calToSearch)
        {
            String alternateLink = ""; ;

            foreach (AtomLink link in _calToSearch.Links)
            {
                // Check all the Links within the CalendarEntry and pick out the alternate to use
                // Alternate links MUST be used for searches to work properly
                if (link.Rel.ToString().Equals("alternate"))
                {
                    alternateLink = link.HRef.ToString();
                }
            }

            if (bUseSSL)
            {
                alternateLink.Replace("http://", "https://");
            }

            return alternateLink;
        }

        /// <summary>
        /// Method used to insert a calendar entry into the specified calendar.
        /// </summary>
        /// <param name="_toAdd">Event Entry to add to calendar</param>
        /// <param name="_calendarName">Calendar name to insert event into</param>
        public EventEntry InsertEventTo(String _calendarName, EventEntry _toAdd)
        {
            EventEntry insertedEvent;
            CalendarEntry cal = GetCalendarByName(_calendarName);
            Reminder rem;

            if (bNotifications)
            {
                MessageBox.Show("yes", "yes", MessageBoxButtons.OK);
                rem = new Reminder();
                rem.Minutes = 30;
                rem.Method = Reminder.ReminderMethod.email;
                _toAdd.Reminders.Add(rem);
            }

            try
            {
                insertedEvent = (EventEntry)csService.Insert(new Uri(GetAlternateURL(cal)), _toAdd);
            }
            catch (Exception e)
            {
                return (EventEntry)null;
            }

            return insertedEvent;
        }

        /// <summary>
        /// Method used to insert a calendar entry into the specified calendar.
        /// </summary>
        /// <param name="_toAdd">Event Entry to add to calendar</param>
        public EventEntry InsertEvent(EventEntry _toAdd)
        {
            Reminder rem;
            EventEntry insertedEvent;
            CalendarEntry cal = GetCalendarByName(defaultCalendar);

            if (bNotifications)
            {                
                rem = new Reminder();
                rem.Minutes = 30;
                rem.Method = Reminder.ReminderMethod.alert;
                _toAdd.Reminders.Add(rem);
            }

            try
            {
                insertedEvent = (EventEntry)csService.Insert(new Uri(GetAlternateURL(cal)), _toAdd);
            }
            catch (Exception e)
            {
                return (EventEntry)null;
            }

            return insertedEvent;
        }

        /// <summary>
        /// Accessor method for Google Login
        /// </summary>
        public String Login
        {
            get {
                return sGoogleLogin;
            }
            set
            {
                sGoogleLogin = value;
                csService.setUserCredentials(sGoogleLogin, sGooglePassword); 
            }
        }

        /// <summary>
        /// Accessor method for Google password
        /// </summary>
        public String Password
        {
            get
            {
                return sGooglePassword;
            }
            set
            {
                sGooglePassword = value;
                csService.setUserCredentials(sGoogleLogin, sGooglePassword); 
            }
        }

        /// <summary>
        /// Accessor method for Google password
        /// </summary>
        public String Calendar
        {
            get
            {
                return defaultCalendar;
            }
            set
            {
                defaultCalendar = value;
            }
        }

        /// <summary>
        /// Accessor method for UseSSl
        /// </summary>
        public Boolean UseSSL
        {
            get
            {
                return bUseSSL;
            }
            set
            {
                bUseSSL = value;
                if (bUseSSL) { sOwnerCalURI.Replace("http://", "https://"); }
                else { sOwnerCalURI.Replace("https://", "http://"); }
            }
        }

        /// <summary>
        /// Accessor method for Notifications
        /// </summary>
        public Boolean Notifications
        {
            get
            {
                return bNotifications;
            }
            set
            {
                bNotifications = value;
            }
        }

        /// <summary>
        /// Accessor method for Days ahead information
        /// </summary>
        public int DaysAhead
        {
            get
            {
                return iDaysAhead;
            }
            set
            {
                iDaysAhead = value;
            }
        }


        // Class variables
        CalendarService csService;
        private String sHttpProtocol;        
        private String sOwnerCalURI;
        private String sGoogleLogin;
        private String sGooglePassword;
        private Boolean bUseSSL;
        private Boolean bNotifications;
        private DateTime dtStartDate, dtEndDate;
        private string defaultCalendar;
        private int iDaysAhead;
    }
}
