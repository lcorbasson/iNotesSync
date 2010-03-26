using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NotesToGoogle
{
    /// <summary>
    /// The NotesServiceConnect class is used to connect to the Lotus notes web server.  Unsure how many
    /// different XML requests can be made, but will look into doing more than just Calendar Displays.
    /// </summary>
    class NotesServiceConnect
    {
        /// <summary>
        /// Constructor method for NotesServiceConnect
        /// </summary>
        /// <param name="_NotesLogin">Username for Lotus notes</param>
        /// <param name="_NotesPassword">Password for Lotus Notes</param>
        /// <param name="_NotesUrl">URL for iNotes webmail</param>
        /// <param name="_ServerAuth">Flag whether to use server auth (or lotus auth)</param>
        /// <param name="_DaysAhead">Number of days ahead to check calendar</param>
        public NotesServiceConnect(String _NotesLogin, String _NotesPassword, String _NotesUrl, 
                Boolean _ServerAuth, int _DaysAhead)
        {
            ncLotusCred = new NetworkCredential(_NotesLogin, _NotesPassword);
            sNotesUrl = _NotesUrl;
            bServerAuth = _ServerAuth;
            iDaysAhead = _DaysAhead;

            if (!bServerAuth)
            {
                // Implement later, have to Login using Notes Auth
            }
        }

        /// <summary>
        /// Constructor method for NotesServiceConnect
        /// </summary>
        /// <param name="_NotesLogin">Username for Lotus notes</param>
        /// <param name="_NotesPassword">Password for Lotus Notes</param>
        /// <param name="_NotesUrl">URL for iNotes webmail</param>
        /// <param name="_ServerAuth">Flag whether to use server auth (or lotus auth)</param>
        public NotesServiceConnect(String _NotesLogin, String _NotesPassword, String _NotesUrl, 
            Boolean _ServerAuth)
        {
            ncLotusCred = new NetworkCredential(_NotesLogin, _NotesPassword);
            sNotesUrl = _NotesUrl;
            bServerAuth = _ServerAuth;
            iDaysAhead = 14;

            if (!bServerAuth)
            {
                // Implement later, have to Login using Notes Auth
            }
        }

        /// <summary>
        /// Constructor method for NotesServiceConnect
        /// </summary>
        /// <param name="_NotesLogin">Username for Lotus notes</param>
        /// <param name="_NotesPassword">Password for Lotus Notes</param>
        /// <param name="_NotesUrl">URL for iNotes webmail</param>
        /// <param name="_DaysAhead">Number of Days ahead to search</param>
        public NotesServiceConnect(String _NotesLogin, String _NotesPassword, String _NotesUrl,
            int _DaysAhead)
        {
            ncLotusCred = new NetworkCredential(_NotesLogin, _NotesPassword);
            sNotesUrl = _NotesUrl;
            bServerAuth = true;
            iDaysAhead = _DaysAhead;

            if (!bServerAuth)
            {
                // Implement later, have to Login using Notes Auth
            }
        }

        /// <summary>
        /// Constructor method for NotesServiceConnect
        /// </summary>
        /// <param name="_NotesLogin">Username for Lotus notes</param>
        /// <param name="_NotesPassword">Password for Lotus Notes</param>
        /// <param name="_NotesUrl">URL for iNotes webmail</param>
        public NotesServiceConnect(String _NotesLogin, String _NotesPassword, String _NotesUrl)
        {
            ncLotusCred = new NetworkCredential(_NotesLogin, _NotesPassword);
            sNotesUrl = _NotesUrl;
            bServerAuth = true;
            iDaysAhead = 14;

            if (!bServerAuth)
            {
                // Implement later, have to Login using Notes Auth
            }
        }

        /// <summary>
        /// Method used to release the connection
        /// </summary>
        public void CloseConnection()
        {
            resStream.Close();
            xmlResponse.Close();
            xmlRequest.Abort();
        }

        /// <summary>
        /// Method used to connect to lotus notes and gather the calendar information entries for the configured
        /// URL / request path.
        /// </summary>
        /// <returns>A stream holding the contents of the HttpWebResponse</returns>
        public StreamReader GetNotesXMLCalendar()
        {
            String iNotesUrl;           

            DateTime startDateTime = DateTime.Now;
            DateTime endDateTime = startDateTime.AddDays(iDaysAhead);

            String startString = startDateTime.ToString("yyyyMMddT000000");
            String endString = endDateTime.ToString("yyyyMMddT235959");

            // Setup/Build the Lotus Notes request URL
            iNotesUrl = sNotesUrl + "/($calendar)?ReadViewEntries&KeyType=time&StartKey=" + startString + "&UntilKey=" + endString;
            //MessageBox.Show(iNotesUrl, "URL", MessageBoxButtons.OK);
            xmlRequest = (HttpWebRequest)WebRequest.Create(iNotesUrl);
            xmlRequest.Credentials = ncLotusCred;
            xmlRequest.KeepAlive = false;
            xmlRequest.Timeout = 20000;

            try
            {               
                // Call the request
                xmlResponse = (HttpWebResponse)xmlRequest.GetResponse();

                // Work with response/Return data
                resStream = new StreamReader(xmlResponse.GetResponseStream());
            }
            catch (WebException e)
            {
                MessageBox.Show(e.Message, "error", MessageBoxButtons.OK);                
                resStream.Close();
                xmlResponse.Close();
                return (StreamReader)null;
            }

            return resStream;
        }

        /// <summary>
        /// Test the connection to the notes server
        /// </summary>
        /// <returns>Whether or not the connection is ok</returns>
        public Boolean TestConnection()
        {
            HttpWebResponse _webResponse;
            String iNotesUrl;

            try
            {
                // Setup/Build the Lotus Notes request URL
                iNotesUrl = sNotesUrl + "/($calendar)?ReadViewEntries";
                HttpWebRequest _webRequest = (HttpWebRequest)WebRequest.Create(iNotesUrl);
                _webRequest.Credentials = ncLotusCred;
                _webResponse = (HttpWebResponse)_webRequest.GetResponse();

                // Close the connections
                _webResponse.Close();
                _webRequest.Abort();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "error", MessageBoxButtons.OK);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Accessor method for sNotesUrl
        /// </summary>
        public String NotesURL
        {
            get
            {
                return sNotesUrl;
            }
            set
            {
                sNotesUrl = value;
            }
        }

        /// <summary>
        /// Accessor method for sNotesLogin
        /// </summary>
        public String Login
        {
            get
            {
                return ncLotusCred.UserName;
            }
            set
            {
                ncLotusCred.UserName = value;
            }
        }

        /// <summary>
        /// Accessor method for sNotesPassword
        /// </summary>
        public String Password
        {
            get
            {
                return ncLotusCred.Password;
            }
            set
            {
                ncLotusCred.Password = value;
            }
        }

        /// <summary>
        /// Accessor method for bServerAuth
        /// </summary>
        public Boolean Auth
        {
            get
            {
                return bServerAuth;
            }
            set
            {
                bServerAuth = value;
            }
        }

        /// <summary>
        /// Accessor method for the Days Ahead variable
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

        // Private class variables
        private String sNotesUrl;
        private NetworkCredential ncLotusCred;
        private Boolean bServerAuth;
        private int iDaysAhead;
        private HttpWebResponse xmlResponse;
        private HttpWebRequest xmlRequest;
        StreamReader resStream;
    }
}
