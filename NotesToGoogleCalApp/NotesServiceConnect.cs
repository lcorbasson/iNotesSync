using System;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

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
            lasterror = "";

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
            lasterror = "";

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
            WebProxy prox;
            RegistryKey regCurrentUser, regUsers;
            String sCurrentUserProxy, sUsersProxy;
            String proxyHost;
            CookieContainer cookies;

            DateTime startDateTime = DateTime.Now;
            DateTime endDateTime = startDateTime.AddDays(iDaysAhead);

            String startString = startDateTime.ToString("yyyyMMddT000000");
            String endString = endDateTime.ToString("yyyyMMddT235959");

            // Setup/Build the Lotus Notes request URL
            iNotesUrl = sNotesUrl + "/($calendar)?ReadViewEntries&KeyType=time&StartKey=" + startString + "&UntilKey=" + endString;            
            
            // If the notes URL contains HTTPS, we need to invalidate the cert, this assumes we know we are connecting to
            // a work server and will always be true
            if (iNotesUrl.Contains("https://"))
            {   // http://support.microsoft.com/kb/823177/en-us
                ServicePointManager.ServerCertificateValidationCallback =
                    new RemoteCertificateValidationCallback(
                    delegate(
                        object sender2,
                        X509Certificate certificate,
                        X509Chain chain,
                        SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    });
            }

            try
            {
                cookies = new CookieContainer();
                xmlRequest = (HttpWebRequest)WebRequest.Create(iNotesUrl);
                xmlRequest.CookieContainer = cookies;
                xmlRequest.Accept = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, application/x-shockwave-flash, application/x-silverlight, */*";
                xmlRequest.UserAgent = "iNotes to Google";
                xmlRequest.KeepAlive = false;
                xmlRequest.Method = "POST";

                // Differenciate between (?Login) and Server Authentication
                if (bServerAuth)
                {
                    //loginRequest = (HttpWebRequest)WebRequest.Create(sNotesUrl + "/?Login&username="+ncLotusCred.UserName.ToString()+"&password="+ncLotusCred.Password.ToString());
                    loginRequest = (HttpWebRequest)WebRequest.Create(sNotesUrl + "/?Login");
                    loginRequest.CookieContainer = cookies;
                    loginRequest.Accept = "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, application/x-shockwave-flash, application/x-silverlight, */*";
                    loginRequest.UserAgent = "iNotes to Google";
                    loginRequest.Method = "POST";
                    loginRequest.KeepAlive = true;
                    loginRequest.MaximumAutomaticRedirections = 5;
                    loginRequest.AllowAutoRedirect = true;

                    // Add login data to POST
                    var requestStream = loginRequest.GetRequestStream();
                    var loginData = Encoding.Default.GetBytes("&username=" + ncLotusCred.UserName.ToString() + "&password=" + ncLotusCred.Password.ToString());
                    requestStream.Write(loginData, 0, loginData.Length);
                    requestStream.Close();

                    // Login
                    loginResponse = loginRequest.GetResponse() as HttpWebResponse;

                    //StreamReader reader2 = new StreamReader(loginResponse.GetResponseStream(), System.Text.Encoding.UTF8);
                    //MessageBox.Show(reader2.ReadToEnd(), "stream", MessageBoxButtons.OK);
                }
                else
                {
                    xmlRequest.Credentials = ncLotusCred;
                }
                                 
                // Call the request
                xmlResponse = xmlRequest.GetResponse() as HttpWebResponse;

                // Work with response/Return data
                resStream = new StreamReader(xmlResponse.GetResponseStream());
                //MessageBox.Show(resStream.ReadToEnd(), "stream", MessageBoxButtons.OK);
            }
            catch (WebException e)
            {
                this.lasterror = "Getting Notes Data as XML: " + e.Message;
                //MessageBox.Show(e.Message, "error", MessageBoxButtons.OK);                
                if (resStream != null) { resStream.Close(); }
                if (xmlResponse != null) { xmlResponse.Close(); }
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
            //HttpWebResponse _webResponse;
            String iNotesUrl;
            WebProxy prox;
            RegistryKey regCurrentUser, regUsers;
            String sCurrentUserProxy, sUsersProxy;
            String proxyHost;

            try
            {
                // Setup/Build the Lotus Notes request URL
                iNotesUrl = sNotesUrl + "/($calendar)?ReadViewEntries";
                HttpWebRequest _webRequest = (HttpWebRequest)WebRequest.Create(iNotesUrl);
                _webRequest.Credentials = ncLotusCred;
                _webRequest.KeepAlive = false;
                _webRequest.Method = "POST";
                _webRequest.Timeout = 20000;

                /*
                // If there is proxy information in registry
                regCurrentUser = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings");
                regUsers = Registry.Users.OpenSubKey(".DEFAULT\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings");

                prox = new WebProxy();

                if ((regCurrentUser != null) && Convert.ToInt32(regCurrentUser.GetValue("ProxyEnable").ToString()) == 1)
                {
                    sCurrentUserProxy = regCurrentUser.GetValue("ProxyServer").ToString();
                    if (sCurrentUserProxy.Contains("http=")) {
                        int httploc = sCurrentUserProxy.IndexOf("http=");
                        proxyHost = sCurrentUserProxy.Substring(httploc + 5, sCurrentUserProxy.IndexOf(";", httploc)-1);
                        prox.Address = new Uri(proxyHost);
                    } else {
                        proxyHost = "http://" + sCurrentUserProxy;
                        prox.Address = new Uri(proxyHost);
                    }
                }

                if ((regUsers != null) && Convert.ToInt32(regUsers.GetValue("ProxyEnable").ToString()) == 1)
                {
                    sUsersProxy = regUsers.GetValue("ProxyServer").ToString();
                    if (sUsersProxy.Contains("http="))
                    {
                        int httploc = sUsersProxy.IndexOf("http=");
                        proxyHost = sUsersProxy.Substring(httploc + 5, sUsersProxy.IndexOf(";", httploc) - 1);
                        prox.Address = new Uri(proxyHost);
                    }
                    else
                    {
                        proxyHost = "http://" + sUsersProxy;
                        prox.Address = new Uri(proxyHost);
                    }
                }

                _webRequest.Proxy = prox;
                */

                // If the notes URL contains HTTPS, we need to invalidate the cert, this assumes we know we are connecting to
                // a work server and will always be true
                if (iNotesUrl.Contains("https://"))
                {   // http://support.microsoft.com/kb/823177/en-us
                    ServicePointManager.ServerCertificateValidationCallback =
                        new RemoteCertificateValidationCallback(
                        delegate(
                            object sender2,
                            X509Certificate certificate,
                            X509Chain chain,
                            SslPolicyErrors sslPolicyErrors)
                        {
                        return true;
                    });
                }

                //_webResponse = (HttpWebResponse)_webRequest.GetResponse();

                // Close the connections
                //_webResponse.Close();
                _webRequest.Abort();
            }
            catch (Exception e)
            {
                this.lasterror = "Testing Notes Connection: " + e.Message;
                //MessageBox.Show(e.Message, "error", MessageBoxButtons.OK);
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

        /// <summary>
        /// set get for last error
        /// </summary>
        public string LastError
        {
            get
            {
                return lasterror;
            }
        }

        // Private class variables
        private String sNotesUrl, sLoginData;
        private NetworkCredential ncLotusCred;
        private Boolean bServerAuth;
        private int iDaysAhead;
        private HttpWebResponse xmlResponse, loginResponse;
        private HttpWebRequest xmlRequest, loginRequest;
        StreamReader resStream;
        private String lasterror;
    }
}
