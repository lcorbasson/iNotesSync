using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Timers;
using System.Windows.Forms;
using System.Threading;
using System.Resources;
using System.Globalization;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Calendar;
using Google.GData.AccessControl;
using NotesToGoogle;

namespace NotesToGoogle
{
    /// <summary>
    /// Public class NotesToGoogleForm creates the main window used for application.
    /// It is also an entry point for the appliction.
    /// </summary>
    public class NotesToGoogleForm : Form
    {
        /// <summary>
        ///  Constructor for NotesToGoogleForm class
        ///  Creates initializes main window components
        /// </summary>
        public NotesToGoogleForm()
        {
            // Get the config object, contains all properties that have been saved.
            config = new SyncPreferences();
            
            // Setup the Timer variable


            // Setup components
            InitializeComponent();      // Auto-gen GUI
            LoadConfigPreferences();
            SetupAndStartScheduleSync();

            if (checkBox_StartMinimized.Checked)
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.ShowInTaskbar = true;
            }
        }        

        /// <summary>
        /// Method called to setup the ScheduleSync timer with appropriate values; if the user has checked 
        /// to start the scheduled sync the system will:
        /// - Disable the Manual Sync button and start the timer
        /// </summary>
        private void SetupAndStartScheduleSync()
        {
            try
            {                
                timer_Schedule.Stop();
                timer_Schedule.Interval = (Convert.ToInt32(textBox_ScheduleSync.Text) * 60 * 1000);
                //MessageBox.Show(timer_Schedule.Interval.ToString(), "interval", MessageBoxButtons.OK);
                //timer_Schedule.Interval = 10000;
                timer_Schedule.Tick += (Timer_ScheduleSync_Tick);

                if (checkBox_ScheduleSync.Checked)
                {
                    timer_Schedule.Start();
                }
            }
            catch (InvalidCastException e)
            {
                textBox_ScheduleSync.Text = (config.GetPreference("ScheduleTime") != "") ? config.GetPreference("ScheduleTime") : "30";
                PrintStringToDebug(e.Message);
            }
        }

        /// <summary>
        /// Method called to load the config preferences into the Form, called AFTER
        /// the InitilaizeComponents method.
        /// </summary>
        private void LoadConfigPreferences()
        {
            // Load preferences from File
            if (config.LoadPreferences())
            {
                // Update GUI with preference values
                // Save Connection Variables
                textBox_WebmailURL.Text = (config.GetPreference("WebmailURL") != "") ? config.GetPreference("WebmailURL") : "http://webmail.domain.com/mail/name.nsf";
                textBox_NotesLogin.Text = config.GetPreference("NotesLogin");
                textBox_NotesPassword.Text = config.GetPreference("NotesPassword");
                textBox_GoogleLogin.Text = config.GetPreference("GoogleLogin");
                textBox_GooglePassword.Text = config.GetPreference("GooglePassword");

                // Save Preference Variables
                checkBox_ConnectUsingSSL.Checked = (config.GetPreference("ConnectUsingSSL") != "") ? Convert.ToBoolean(config.GetPreference("ConnectUsingSSL")) : true;
                checkBox_NotesServerAuth.Checked = (config.GetPreference("NotesServerAuth") != "") ? Convert.ToBoolean(config.GetPreference("NotesServerAuth")) : true;
                checkBox_MinimizeToTray.Checked = (config.GetPreference("MinimizeToTray") != "") ? Convert.ToBoolean(config.GetPreference("MinimizeToTray")) : false;
                radioButton_MainCalChoice.Checked = (config.GetPreference("MainCalChoice") != "") ? Convert.ToBoolean(config.GetPreference("MainCalChoice")) : true;
                radioButton_OtherCalChoice.Checked = (config.GetPreference("OtherCalChoice") != "") ? Convert.ToBoolean(config.GetPreference("OtherCalChoice")) : false;
                checkBox_CustomDaysAhead.Checked = (config.GetPreference("CustomDaysAhead") != "") ? Convert.ToBoolean(config.GetPreference("CustomDaysAhead")) : false;
                checkBox_CreateNotification.Checked = (config.GetPreference("NotificationsOn") != "") ? Convert.ToBoolean(config.GetPreference("NotificationsOn")) : false;
                textBox_CustomDaysAhead.Text = (config.GetPreference("DaysAhead") != "") ? config.GetPreference("DaysAhead") : "14";
                textBox_OtherCalName.Text = (config.GetPreference("OtherCalName") != "") ? config.GetPreference("OtherCalName") : "Lotus.Notes";
                checkBox_StartMinimized.Checked = (config.GetPreference("StartMinimized") != "") ? Convert.ToBoolean(config.GetPreference("StartMinimized")) : false;

                // Save Schedule Variables
                checkBox_SyncOnStartup.Checked = (config.GetPreference("SyncOnStartup") != "") ? Convert.ToBoolean(config.GetPreference("SyncOnStartup")) : false;
                checkBox_ScheduleSync.Checked = (config.GetPreference("ScheduleSync") != "") ? Convert.ToBoolean(config.GetPreference("ScheduleSync")) : false;
                textBox_ScheduleSync.Text = (config.GetPreference("ScheduleTime") != "") ? config.GetPreference("ScheduleTime") : "30";
            }
            else
            {   // If there were errors, we need to post an error
                PrintStringToDebug("Config Preferences did not load properly");
            }
        }

        /// <summary>
        /// Update method called for changing localization information
        /// </summary>
        private void UpdateUI()
        {
            // Setup information section
            label_DateValue.Text = DateTime.Today.ToString();
            label_VersionValue.Text = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            textBox_AboutValue.Text = "Sync your Lotus Notes calendar to Google.  This allows update to your iPhone and other applications.";
        
            // Fill Top page with culture settings
            this.Text = m_ResourceManager.GetString("Title");
            label_Author.Text = m_ResourceManager.GetString("label_Author");
            label_Version.Text = m_ResourceManager.GetString("label_Version");
            label_Date.Text = m_ResourceManager.GetString("label_Date");
            label_About.Text = m_ResourceManager.GetString("label_About");
            textBox_AboutValue.Text = m_ResourceManager.GetString("textBox_AboutValue");

            // Fill Tab titles with culture settings
            tabPage_Connect.Text = m_ResourceManager.GetString("tabPage_Connect");
            tabPage_Preferences.Text = m_ResourceManager.GetString("tabPage_Preferences");
            tabPage_SchedulingSync.Text = m_ResourceManager.GetString("tabPage_SchedulingSync");
            tabPage_Debug.Text = m_ResourceManager.GetString("tabPage_Debug");

            // Fill Connection with Culture
            label_NotesLogin.Text = m_ResourceManager.GetString("label_NotesLogin");
            label_NotesPassword.Text = m_ResourceManager.GetString("label_NotesPassword");
            label_GoogleLogin.Text = m_ResourceManager.GetString("label_GoogleLogin");
            label_GooglePassword.Text = m_ResourceManager.GetString("label_GooglePassword");

            // Fill Prefernces with Culture
            radioButton_MainCalChoice.Text = m_ResourceManager.GetString("radioButton_MainCalChoice");
            radioButton_OtherCalChoice.Text = m_ResourceManager.GetString("radioButton_OtherCalChoice");
            checkBox_ConnectUsingSSL.Text = m_ResourceManager.GetString("checkBox_ConnectUsingSSL");
            checkBox_CreateNotification.Text = m_ResourceManager.GetString("checkBox_CreateNotification");
            checkBox_NotesServerAuth.Text = m_ResourceManager.GetString("checkBox_NotesServerAuth");
            checkBox_CustomDaysAhead.Text = m_ResourceManager.GetString("checkBox_CustomDaysAhead");
            checkBox_MinimizeToTray.Text = m_ResourceManager.GetString("checkBox_MinimizeToTray");
            checkBox_StartMinimized.Text = m_ResourceManager.GetString("checkBox_StartMinimized");
            button_CreateService.Text = m_ResourceManager.GetString("button_CreateService");
            radioButton_English.Text = m_ResourceManager.GetString("radioButton_English");
            radioButton_French.Text = m_ResourceManager.GetString("radioButton_French");

            // Fill Schedulesync with culture
            checkBox_SyncOnStartup.Text = m_ResourceManager.GetString("checkBox_SyncOnStartup");
            checkBox_ScheduleSync.Text = m_ResourceManager.GetString("checkBox_ScheduleSync");
            
            // Fill Debug with culture
            button_ClearDebug.Text = m_ResourceManager.GetString("button_ClearDebug");
            checkBox_MessageBoxDebug.Text = m_ResourceManager.GetString("checkBox_MessageBoxDebug");

            // Fill Bottom section with Culture
            button_SaveSettings.Text = m_ResourceManager.GetString("button_SaveSettings");
            button_ManualSync.Text = m_ResourceManager.GetString("button_ManualSync");
            label_Status.Text = m_ResourceManager.GetString("label_Status");
        }

        /// <summary>
        /// Override method from parent.OnResize; method will move application to tray
        /// </summary>
        /// <param name="e">Event args passed to on resize</param>
        protected override void OnResize(EventArgs e)
        {   // Call the parent OnResize
            base.OnResize(e);

            //m_previousWindowState = (this.WindowState != FormWindowState.Minimized) ? FormWindowState.Normal : FormWindowState.Minimized;                        
            if (checkBox_MinimizeToTray.Checked)
            {   // We need to keep track of whether you're minimizing from a normal or maximized window 
                notifyIcon_Tray.Visible = (this.WindowState == FormWindowState.Minimized);
                this.Visible = !notifyIcon_Tray.Visible;
            }

            this.ShowInTaskbar = this.Visible;
        }

        /// <summary>
        /// Method called when the main form is loaded
        /// </summary>
        /// <param name="sender">Object taht sent the Load event</param>
        /// <param name="e">Event Arguments passed</param>
        private void NotesToGoogleForm_Load(object sender, EventArgs e)
        {   /*
            if (config.LoadPreferences())
            {
                // Update GUI with preference values
                // Save Connection Variables
                textBox_WebmailURL.Text = (config.GetPreference("WebmailURL") != "") ? config.GetPreference("WebmailURL") : "http://webmail.domain.com/mail/name.nsf";
                textBox_NotesLogin.Text = config.GetPreference("NotesLogin");
                textBox_NotesPassword.Text = config.GetPreference("NotesPassword");
                textBox_GoogleLogin.Text = config.GetPreference("GoogleLogin");
                textBox_GooglePassword.Text = config.GetPreference("GooglePassword");

                // Save Preference Variables
                checkBox_ConnectUsingSSL.Checked = (config.GetPreference("ConnectUsingSSL") != "") ? Convert.ToBoolean(config.GetPreference("ConnectUsingSSL")) : true;
                checkBox_NotesServerAuth.Checked = (config.GetPreference("NotesServerAuth") != "") ? Convert.ToBoolean(config.GetPreference("NotesServerAuth")) : true;
                checkBox_MinimizeToTray.Checked = (config.GetPreference("MinimizeToTray") != "") ? Convert.ToBoolean(config.GetPreference("MinimizeToTray")) : false;
                radioButton_MainCalChoice.Checked = (config.GetPreference("MainCalChoice") != "") ? Convert.ToBoolean(config.GetPreference("MainCalChoice")) : true;
                radioButton_OtherCalChoice.Checked = (config.GetPreference("OtherCalChoice") != "") ? Convert.ToBoolean(config.GetPreference("OtherCalChoice")) : false;
                checkBox_CustomDaysAhead.Checked = (config.GetPreference("CustomDaysAhead") != "") ? Convert.ToBoolean(config.GetPreference("CustomDaysAhead")) : false;
                checkBox_CreateNotification.Checked = (config.GetPreference("NotificationsOn") != "") ? Convert.ToBoolean(config.GetPreference("NotificationsOn")) : false;
                textBox_CustomDaysAhead.Text = (config.GetPreference("DaysAhead") != "") ? config.GetPreference("DaysAhead") : "14";
                textBox_OtherCalName.Text = (config.GetPreference("OtherCalName") != "") ? config.GetPreference("OtherCalName") : "Lotus.Notes";
                checkBox_StartMinimized.Checked = (config.GetPreference("StartMinimized") != "") ? Convert.ToBoolean(config.GetPreference("StartMinimized")) : false;

                // Save Schedule Variables
                checkBox_SyncOnStartup.Checked = (config.GetPreference("SyncOnStartup") != "") ? Convert.ToBoolean(config.GetPreference("SyncOnStartup")) : false;
                checkBox_ScheduleSync.Checked = (config.GetPreference("ScheduleSync") != "") ? Convert.ToBoolean(config.GetPreference("ScheduleSync")) : false;
                textBox_ScheduleSync.Text = (config.GetPreference("ScheduleTime") != "") ? config.GetPreference("ScheduleTime") : "30";
            }
            else
            {   // If there were errors, we need to post an error
                PrintStringToDebug("Config Preferences did not load properly");
            }*/

            // Setup
            Thread.CurrentThread.CurrentUICulture = m_EnglishCulture;
            UpdateUI();                 // Personal Components

            // Sync on Startup if selected
            if (checkBox_SyncOnStartup.Checked)
            {   // If the user has checked to Sync on Startup, we simulate the Button click
                button_ManualSync_Click((object)null, (EventArgs)null);
            }
        }

        /// <summary>
        /// Override event for when the Form is closing; if a background worker is syncing, we want to warn
        /// the users not to close.
        /// </summary>
        /// <param name="sender">Object that fired the event</param>
        /// <param name="e">Parameters for the cancel event</param>
        private void FormClosingEventCancel_Closing(object sender, CancelEventArgs e)
        {
            DialogResult result;

            if (backgroundWorker_SingleSync.IsBusy)
            {
                result = MessageBox.Show("There is a sync job currently running; are you sure you'd like to quit?", "Job running", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes) 
                {   // Forcably kill the background process and quit
                    backgroundWorker_SingleSync.Dispose();
                    e.Cancel = false; 
                }
                else 
                {   // Cancel quit
                    e.Cancel = true; 
                }
            }
            else
            {
                e.Cancel = false;
            }
        }

        /// <summary>
        /// Override method for disposing the Form Components
        /// </summary>
        /// <param name="disposing">Flag of weather disposing</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Method used to display the XML content in the Debug panel.  This should be called by a thread
        /// safe delegate in the DoTheWork thread method.
        /// </summary>
        /// <param name="stream"></param>
        private void PrintXMLtoDebug(Stream stream)
        {
            DateTime currentDebug = DateTime.Now;

            this.textBox_Debug.Text += " " + currentDebug.ToString() + "\r\n";

            StreamReader reader2 = new StreamReader(stream, System.Text.Encoding.UTF8);
            this.textBox_Debug.Text += reader2.ReadToEnd();
            this.textBox_Debug.Text += "\r\n \r\n";

            this.textBox_Debug.Select(textBox_Debug.Text.Length, 0);
            this.textBox_Debug.ScrollToCaret();
        }

        /// <summary>
        /// Method used to display error messages in the debug log. This should be called by a thread
        /// safe delegate in the DoTheWork thread method.
        /// </summary>
        /// <param name="message">Message to be added to the debug log</param>
        private void PrintStringToDebug(String message)
        {
            DateTime currentDebug = DateTime.Now;

            this.textBox_Debug.Text += " " + currentDebug.ToString() + "\r\n";
            this.textBox_Debug.Text += message + "\r\n \r\n";

            this.textBox_Debug.Select(textBox_Debug.Text.Length, 0);
            this.textBox_Debug.ScrollToCaret();

            if (checkBox_MessageBoxDebug.Checked)
            {
                MessageBox.Show(message, "Message", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new NotesToGoogleForm());
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NotesToGoogleForm));
            this.tabControl_MainControl = new System.Windows.Forms.TabControl();
            this.tabPage_Connect = new System.Windows.Forms.TabPage();
            this.groupBox_GoogleConnection = new System.Windows.Forms.GroupBox();
            this.label_GoogleLogin = new System.Windows.Forms.Label();
            this.label_GooglePassword = new System.Windows.Forms.Label();
            this.textBox_GooglePassword = new System.Windows.Forms.TextBox();
            this.textBox_GoogleLogin = new System.Windows.Forms.TextBox();
            this.groupBox_LotusConnection = new System.Windows.Forms.GroupBox();
            this.label_NotesWebmailURL = new System.Windows.Forms.Label();
            this.textBox_WebmailURL = new System.Windows.Forms.TextBox();
            this.label_NotesLogin = new System.Windows.Forms.Label();
            this.textBox_NotesPassword = new System.Windows.Forms.TextBox();
            this.textBox_NotesLogin = new System.Windows.Forms.TextBox();
            this.label_NotesPassword = new System.Windows.Forms.Label();
            this.tabPage_Preferences = new System.Windows.Forms.TabPage();
            this.groupBox_language = new System.Windows.Forms.GroupBox();
            this.radioButton_English = new System.Windows.Forms.RadioButton();
            this.radioButton_French = new System.Windows.Forms.RadioButton();
            this.groupBox_Window = new System.Windows.Forms.GroupBox();
            this.checkBox_StartMinimized = new System.Windows.Forms.CheckBox();
            this.checkBox_MinimizeToTray = new System.Windows.Forms.CheckBox();
            this.button_CreateService = new System.Windows.Forms.Button();
            this.groupBox_LotusNotes = new System.Windows.Forms.GroupBox();
            this.checkBox_NotesServerAuth = new System.Windows.Forms.CheckBox();
            this.checkBox_CustomDaysAhead = new System.Windows.Forms.CheckBox();
            this.textBox_CustomDaysAhead = new System.Windows.Forms.TextBox();
            this.groupBox_GoogleCalendar = new System.Windows.Forms.GroupBox();
            this.radioButton_MainCalChoice = new System.Windows.Forms.RadioButton();
            this.radioButton_OtherCalChoice = new System.Windows.Forms.RadioButton();
            this.textBox_OtherCalName = new System.Windows.Forms.TextBox();
            this.checkBox_ConnectUsingSSL = new System.Windows.Forms.CheckBox();
            this.checkBox_CreateNotification = new System.Windows.Forms.CheckBox();
            this.tabPage_SchedulingSync = new System.Windows.Forms.TabPage();
            this.groupBox_Sync = new System.Windows.Forms.GroupBox();
            this.checkBox_ScheduleSync = new System.Windows.Forms.CheckBox();
            this.label_Minutes = new System.Windows.Forms.Label();
            this.checkBox_SyncOnStartup = new System.Windows.Forms.CheckBox();
            this.textBox_ScheduleSync = new System.Windows.Forms.TextBox();
            this.tabPage_Debug = new System.Windows.Forms.TabPage();
            this.checkBox_MessageBoxDebug = new System.Windows.Forms.CheckBox();
            this.button_ClearDebug = new System.Windows.Forms.Button();
            this.textBox_Debug = new System.Windows.Forms.TextBox();
            this.label_Author = new System.Windows.Forms.Label();
            this.label_Version = new System.Windows.Forms.Label();
            this.label_Date = new System.Windows.Forms.Label();
            this.label_VersionValue = new System.Windows.Forms.Label();
            this.label_DateValue = new System.Windows.Forms.Label();
            this.label_About = new System.Windows.Forms.Label();
            this.textBox_AboutValue = new System.Windows.Forms.TextBox();
            this.button_ManualSync = new System.Windows.Forms.Button();
            this.button_SaveSettings = new System.Windows.Forms.Button();
            this.panel_StatusPanel = new System.Windows.Forms.Panel();
            this.progressBar_SyncProgress = new System.Windows.Forms.ProgressBar();
            this.label_Status = new System.Windows.Forms.Label();
            this.pictureBox_Logo = new System.Windows.Forms.PictureBox();
            this.backgroundWorker_SingleSync = new System.ComponentModel.BackgroundWorker();
            this.notifyIcon_Tray = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextMenu_Tray = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem_Open = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem_Sync = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem_Close = new System.Windows.Forms.ToolStripMenuItem();
            this.linkLabel_url = new System.Windows.Forms.LinkLabel();
            this.backgroundWorker_Schedule = new System.ComponentModel.BackgroundWorker();
            this.timer_Schedule = new System.Windows.Forms.Timer(this.components);
            this.tabControl_MainControl.SuspendLayout();
            this.tabPage_Connect.SuspendLayout();
            this.groupBox_GoogleConnection.SuspendLayout();
            this.groupBox_LotusConnection.SuspendLayout();
            this.tabPage_Preferences.SuspendLayout();
            this.groupBox_language.SuspendLayout();
            this.groupBox_Window.SuspendLayout();
            this.groupBox_LotusNotes.SuspendLayout();
            this.groupBox_GoogleCalendar.SuspendLayout();
            this.tabPage_SchedulingSync.SuspendLayout();
            this.groupBox_Sync.SuspendLayout();
            this.tabPage_Debug.SuspendLayout();
            this.panel_StatusPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Logo)).BeginInit();
            this.contextMenu_Tray.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl_MainControl
            // 
            this.tabControl_MainControl.Controls.Add(this.tabPage_Connect);
            this.tabControl_MainControl.Controls.Add(this.tabPage_Preferences);
            this.tabControl_MainControl.Controls.Add(this.tabPage_SchedulingSync);
            this.tabControl_MainControl.Controls.Add(this.tabPage_Debug);
            this.tabControl_MainControl.Location = new System.Drawing.Point(10, 139);
            this.tabControl_MainControl.Multiline = true;
            this.tabControl_MainControl.Name = "tabControl_MainControl";
            this.tabControl_MainControl.SelectedIndex = 0;
            this.tabControl_MainControl.Size = new System.Drawing.Size(580, 262);
            this.tabControl_MainControl.TabIndex = 0;
            // 
            // tabPage_Connect
            // 
            this.tabPage_Connect.BackColor = System.Drawing.Color.Transparent;
            this.tabPage_Connect.Controls.Add(this.groupBox_GoogleConnection);
            this.tabPage_Connect.Controls.Add(this.groupBox_LotusConnection);
            this.tabPage_Connect.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.tabPage_Connect.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Connect.Name = "tabPage_Connect";
            this.tabPage_Connect.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Connect.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.tabPage_Connect.Size = new System.Drawing.Size(572, 236);
            this.tabPage_Connect.TabIndex = 0;
            this.tabPage_Connect.Text = "Connection";
            this.tabPage_Connect.ToolTipText = "Set Connection Information";
            // 
            // groupBox_GoogleConnection
            // 
            this.groupBox_GoogleConnection.Controls.Add(this.label_GoogleLogin);
            this.groupBox_GoogleConnection.Controls.Add(this.label_GooglePassword);
            this.groupBox_GoogleConnection.Controls.Add(this.textBox_GooglePassword);
            this.groupBox_GoogleConnection.Controls.Add(this.textBox_GoogleLogin);
            this.groupBox_GoogleConnection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_GoogleConnection.Location = new System.Drawing.Point(14, 134);
            this.groupBox_GoogleConnection.Name = "groupBox_GoogleConnection";
            this.groupBox_GoogleConnection.Size = new System.Drawing.Size(537, 91);
            this.groupBox_GoogleConnection.TabIndex = 13;
            this.groupBox_GoogleConnection.TabStop = false;
            this.groupBox_GoogleConnection.Text = " Google ";
            // 
            // label_GoogleLogin
            // 
            this.label_GoogleLogin.AutoSize = true;
            this.label_GoogleLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_GoogleLogin.Location = new System.Drawing.Point(6, 24);
            this.label_GoogleLogin.Name = "label_GoogleLogin";
            this.label_GoogleLogin.Size = new System.Drawing.Size(110, 15);
            this.label_GoogleLogin.TabIndex = 5;
            this.label_GoogleLogin.Text = "Username (email):";
            // 
            // label_GooglePassword
            // 
            this.label_GooglePassword.AutoSize = true;
            this.label_GooglePassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_GooglePassword.Location = new System.Drawing.Point(6, 54);
            this.label_GooglePassword.Name = "label_GooglePassword";
            this.label_GooglePassword.Size = new System.Drawing.Size(61, 15);
            this.label_GooglePassword.TabIndex = 6;
            this.label_GooglePassword.Text = "Password";
            // 
            // textBox_GooglePassword
            // 
            this.textBox_GooglePassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_GooglePassword.Location = new System.Drawing.Point(139, 51);
            this.textBox_GooglePassword.MaxLength = 20;
            this.textBox_GooglePassword.Name = "textBox_GooglePassword";
            this.textBox_GooglePassword.PasswordChar = '*';
            this.textBox_GooglePassword.Size = new System.Drawing.Size(177, 21);
            this.textBox_GooglePassword.TabIndex = 11;
            // 
            // textBox_GoogleLogin
            // 
            this.textBox_GoogleLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_GoogleLogin.Location = new System.Drawing.Point(139, 21);
            this.textBox_GoogleLogin.MaxLength = 50;
            this.textBox_GoogleLogin.Name = "textBox_GoogleLogin";
            this.textBox_GoogleLogin.Size = new System.Drawing.Size(177, 21);
            this.textBox_GoogleLogin.TabIndex = 10;
            // 
            // groupBox_LotusConnection
            // 
            this.groupBox_LotusConnection.Controls.Add(this.label_NotesWebmailURL);
            this.groupBox_LotusConnection.Controls.Add(this.textBox_WebmailURL);
            this.groupBox_LotusConnection.Controls.Add(this.label_NotesLogin);
            this.groupBox_LotusConnection.Controls.Add(this.textBox_NotesPassword);
            this.groupBox_LotusConnection.Controls.Add(this.textBox_NotesLogin);
            this.groupBox_LotusConnection.Controls.Add(this.label_NotesPassword);
            this.groupBox_LotusConnection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_LotusConnection.Location = new System.Drawing.Point(14, 13);
            this.groupBox_LotusConnection.Name = "groupBox_LotusConnection";
            this.groupBox_LotusConnection.Size = new System.Drawing.Size(537, 111);
            this.groupBox_LotusConnection.TabIndex = 12;
            this.groupBox_LotusConnection.TabStop = false;
            this.groupBox_LotusConnection.Text = " Lotus Notes ";
            // 
            // label_NotesWebmailURL
            // 
            this.label_NotesWebmailURL.AutoSize = true;
            this.label_NotesWebmailURL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_NotesWebmailURL.Location = new System.Drawing.Point(6, 27);
            this.label_NotesWebmailURL.Name = "label_NotesWebmailURL";
            this.label_NotesWebmailURL.Size = new System.Drawing.Size(87, 15);
            this.label_NotesWebmailURL.TabIndex = 2;
            this.label_NotesWebmailURL.Text = "Webmail URL:";
            // 
            // textBox_WebmailURL
            // 
            this.textBox_WebmailURL.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_WebmailURL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_WebmailURL.Location = new System.Drawing.Point(139, 24);
            this.textBox_WebmailURL.MaxLength = 100;
            this.textBox_WebmailURL.Name = "textBox_WebmailURL";
            this.textBox_WebmailURL.Size = new System.Drawing.Size(336, 21);
            this.textBox_WebmailURL.TabIndex = 7;
            // 
            // label_NotesLogin
            // 
            this.label_NotesLogin.AutoSize = true;
            this.label_NotesLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_NotesLogin.Location = new System.Drawing.Point(6, 57);
            this.label_NotesLogin.Name = "label_NotesLogin";
            this.label_NotesLogin.Size = new System.Drawing.Size(41, 15);
            this.label_NotesLogin.TabIndex = 3;
            this.label_NotesLogin.Text = "Login:";
            // 
            // textBox_NotesPassword
            // 
            this.textBox_NotesPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_NotesPassword.Location = new System.Drawing.Point(139, 84);
            this.textBox_NotesPassword.MaxLength = 20;
            this.textBox_NotesPassword.Name = "textBox_NotesPassword";
            this.textBox_NotesPassword.PasswordChar = '*';
            this.textBox_NotesPassword.Size = new System.Drawing.Size(177, 21);
            this.textBox_NotesPassword.TabIndex = 9;
            // 
            // textBox_NotesLogin
            // 
            this.textBox_NotesLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_NotesLogin.Location = new System.Drawing.Point(139, 54);
            this.textBox_NotesLogin.MaxLength = 20;
            this.textBox_NotesLogin.Name = "textBox_NotesLogin";
            this.textBox_NotesLogin.Size = new System.Drawing.Size(177, 21);
            this.textBox_NotesLogin.TabIndex = 8;
            // 
            // label_NotesPassword
            // 
            this.label_NotesPassword.AutoSize = true;
            this.label_NotesPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_NotesPassword.Location = new System.Drawing.Point(6, 87);
            this.label_NotesPassword.Name = "label_NotesPassword";
            this.label_NotesPassword.Size = new System.Drawing.Size(64, 15);
            this.label_NotesPassword.TabIndex = 4;
            this.label_NotesPassword.Text = "Password:";
            // 
            // tabPage_Preferences
            // 
            this.tabPage_Preferences.AutoScroll = true;
            this.tabPage_Preferences.Controls.Add(this.groupBox_language);
            this.tabPage_Preferences.Controls.Add(this.groupBox_Window);
            this.tabPage_Preferences.Controls.Add(this.groupBox_LotusNotes);
            this.tabPage_Preferences.Controls.Add(this.groupBox_GoogleCalendar);
            this.tabPage_Preferences.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.tabPage_Preferences.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Preferences.Name = "tabPage_Preferences";
            this.tabPage_Preferences.Padding = new System.Windows.Forms.Padding(3, 3, 3, 20);
            this.tabPage_Preferences.Size = new System.Drawing.Size(572, 236);
            this.tabPage_Preferences.TabIndex = 1;
            this.tabPage_Preferences.Text = "Preferences";
            this.tabPage_Preferences.UseVisualStyleBackColor = true;
            // 
            // groupBox_language
            // 
            this.groupBox_language.Controls.Add(this.radioButton_English);
            this.groupBox_language.Controls.Add(this.radioButton_French);
            this.groupBox_language.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_language.Location = new System.Drawing.Point(14, 336);
            this.groupBox_language.Name = "groupBox_language";
            this.groupBox_language.Size = new System.Drawing.Size(519, 64);
            this.groupBox_language.TabIndex = 20;
            this.groupBox_language.TabStop = false;
            this.groupBox_language.Text = " Language ";
            // 
            // radioButton_English
            // 
            this.radioButton_English.AutoSize = true;
            this.radioButton_English.Checked = true;
            this.radioButton_English.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_English.Location = new System.Drawing.Point(7, 28);
            this.radioButton_English.Name = "radioButton_English";
            this.radioButton_English.Size = new System.Drawing.Size(66, 19);
            this.radioButton_English.TabIndex = 15;
            this.radioButton_English.TabStop = true;
            this.radioButton_English.Text = "English";
            this.radioButton_English.UseVisualStyleBackColor = true;
            this.radioButton_English.CheckedChanged += new System.EventHandler(this.radioButton_English_CheckedChanged);
            // 
            // radioButton_French
            // 
            this.radioButton_French.AutoSize = true;
            this.radioButton_French.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_French.Location = new System.Drawing.Point(181, 28);
            this.radioButton_French.Name = "radioButton_French";
            this.radioButton_French.Size = new System.Drawing.Size(63, 19);
            this.radioButton_French.TabIndex = 16;
            this.radioButton_French.Text = "French";
            this.radioButton_French.UseVisualStyleBackColor = true;
            this.radioButton_French.CheckedChanged += new System.EventHandler(this.radioButton_French_CheckedChanged);
            // 
            // groupBox_Window
            // 
            this.groupBox_Window.Controls.Add(this.checkBox_StartMinimized);
            this.groupBox_Window.Controls.Add(this.checkBox_MinimizeToTray);
            this.groupBox_Window.Controls.Add(this.button_CreateService);
            this.groupBox_Window.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_Window.Location = new System.Drawing.Point(14, 233);
            this.groupBox_Window.Name = "groupBox_Window";
            this.groupBox_Window.Size = new System.Drawing.Size(519, 94);
            this.groupBox_Window.TabIndex = 19;
            this.groupBox_Window.TabStop = false;
            this.groupBox_Window.Text = "Window";
            // 
            // checkBox_StartMinimized
            // 
            this.checkBox_StartMinimized.AutoSize = true;
            this.checkBox_StartMinimized.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_StartMinimized.Location = new System.Drawing.Point(7, 30);
            this.checkBox_StartMinimized.Name = "checkBox_StartMinimized";
            this.checkBox_StartMinimized.Size = new System.Drawing.Size(112, 19);
            this.checkBox_StartMinimized.TabIndex = 13;
            this.checkBox_StartMinimized.Text = "Start Minimized";
            this.checkBox_StartMinimized.UseVisualStyleBackColor = true;
            // 
            // checkBox_MinimizeToTray
            // 
            this.checkBox_MinimizeToTray.AutoSize = true;
            this.checkBox_MinimizeToTray.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_MinimizeToTray.Location = new System.Drawing.Point(181, 31);
            this.checkBox_MinimizeToTray.Name = "checkBox_MinimizeToTray";
            this.checkBox_MinimizeToTray.Size = new System.Drawing.Size(120, 19);
            this.checkBox_MinimizeToTray.TabIndex = 8;
            this.checkBox_MinimizeToTray.Text = "Minimize To Tray";
            this.checkBox_MinimizeToTray.UseVisualStyleBackColor = true;
            // 
            // button_CreateService
            // 
            this.button_CreateService.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_CreateService.Location = new System.Drawing.Point(6, 59);
            this.button_CreateService.Name = "button_CreateService";
            this.button_CreateService.Size = new System.Drawing.Size(104, 23);
            this.button_CreateService.TabIndex = 11;
            this.button_CreateService.Text = "Create Service";
            this.button_CreateService.UseVisualStyleBackColor = true;
            // 
            // groupBox_LotusNotes
            // 
            this.groupBox_LotusNotes.Controls.Add(this.checkBox_NotesServerAuth);
            this.groupBox_LotusNotes.Controls.Add(this.checkBox_CustomDaysAhead);
            this.groupBox_LotusNotes.Controls.Add(this.textBox_CustomDaysAhead);
            this.groupBox_LotusNotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_LotusNotes.Location = new System.Drawing.Point(14, 135);
            this.groupBox_LotusNotes.Name = "groupBox_LotusNotes";
            this.groupBox_LotusNotes.Size = new System.Drawing.Size(519, 88);
            this.groupBox_LotusNotes.TabIndex = 18;
            this.groupBox_LotusNotes.TabStop = false;
            this.groupBox_LotusNotes.Text = " Lotus Notes ";
            // 
            // checkBox_NotesServerAuth
            // 
            this.checkBox_NotesServerAuth.AutoSize = true;
            this.checkBox_NotesServerAuth.Checked = true;
            this.checkBox_NotesServerAuth.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_NotesServerAuth.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_NotesServerAuth.Location = new System.Drawing.Point(6, 29);
            this.checkBox_NotesServerAuth.Name = "checkBox_NotesServerAuth";
            this.checkBox_NotesServerAuth.Size = new System.Drawing.Size(157, 19);
            this.checkBox_NotesServerAuth.TabIndex = 6;
            this.checkBox_NotesServerAuth.Text = "Use server Athentication";
            this.checkBox_NotesServerAuth.UseVisualStyleBackColor = true;
            // 
            // checkBox_CustomDaysAhead
            // 
            this.checkBox_CustomDaysAhead.AutoSize = true;
            this.checkBox_CustomDaysAhead.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_CustomDaysAhead.Location = new System.Drawing.Point(6, 57);
            this.checkBox_CustomDaysAhead.Name = "checkBox_CustomDaysAhead";
            this.checkBox_CustomDaysAhead.Size = new System.Drawing.Size(212, 19);
            this.checkBox_CustomDaysAhead.TabIndex = 9;
            this.checkBox_CustomDaysAhead.Text = "Configure Days Ahead (default 14)";
            this.checkBox_CustomDaysAhead.UseVisualStyleBackColor = true;
            this.checkBox_CustomDaysAhead.CheckedChanged += new System.EventHandler(this.checkBox_CustomDaysAhead_CheckedChanged);
            // 
            // textBox_CustomDaysAhead
            // 
            this.textBox_CustomDaysAhead.Enabled = false;
            this.textBox_CustomDaysAhead.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_CustomDaysAhead.Location = new System.Drawing.Point(298, 55);
            this.textBox_CustomDaysAhead.Name = "textBox_CustomDaysAhead";
            this.textBox_CustomDaysAhead.Size = new System.Drawing.Size(60, 21);
            this.textBox_CustomDaysAhead.TabIndex = 10;
            this.textBox_CustomDaysAhead.Text = "14";
            this.textBox_CustomDaysAhead.LostFocus += new System.EventHandler(this.textBox_CustomDaysAhead_LostFocus);
            // 
            // groupBox_GoogleCalendar
            // 
            this.groupBox_GoogleCalendar.Controls.Add(this.radioButton_MainCalChoice);
            this.groupBox_GoogleCalendar.Controls.Add(this.radioButton_OtherCalChoice);
            this.groupBox_GoogleCalendar.Controls.Add(this.textBox_OtherCalName);
            this.groupBox_GoogleCalendar.Controls.Add(this.checkBox_ConnectUsingSSL);
            this.groupBox_GoogleCalendar.Controls.Add(this.checkBox_CreateNotification);
            this.groupBox_GoogleCalendar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_GoogleCalendar.Location = new System.Drawing.Point(14, 13);
            this.groupBox_GoogleCalendar.Name = "groupBox_GoogleCalendar";
            this.groupBox_GoogleCalendar.Size = new System.Drawing.Size(519, 109);
            this.groupBox_GoogleCalendar.TabIndex = 17;
            this.groupBox_GoogleCalendar.TabStop = false;
            this.groupBox_GoogleCalendar.Text = " Google Calendar ";
            // 
            // radioButton_MainCalChoice
            // 
            this.radioButton_MainCalChoice.AutoSize = true;
            this.radioButton_MainCalChoice.Checked = true;
            this.radioButton_MainCalChoice.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.radioButton_MainCalChoice.Location = new System.Drawing.Point(6, 27);
            this.radioButton_MainCalChoice.Name = "radioButton_MainCalChoice";
            this.radioButton_MainCalChoice.Size = new System.Drawing.Size(106, 19);
            this.radioButton_MainCalChoice.TabIndex = 1;
            this.radioButton_MainCalChoice.TabStop = true;
            this.radioButton_MainCalChoice.Text = "Main Calendar";
            this.radioButton_MainCalChoice.UseVisualStyleBackColor = true;
            this.radioButton_MainCalChoice.CheckedChanged += new System.EventHandler(this.radioButton_MainCalChoice_CheckedChanged);
            // 
            // radioButton_OtherCalChoice
            // 
            this.radioButton_OtherCalChoice.AutoSize = true;
            this.radioButton_OtherCalChoice.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.radioButton_OtherCalChoice.Location = new System.Drawing.Point(181, 27);
            this.radioButton_OtherCalChoice.Name = "radioButton_OtherCalChoice";
            this.radioButton_OtherCalChoice.Size = new System.Drawing.Size(108, 19);
            this.radioButton_OtherCalChoice.TabIndex = 2;
            this.radioButton_OtherCalChoice.Text = "Other Calendar";
            this.radioButton_OtherCalChoice.UseVisualStyleBackColor = true;
            this.radioButton_OtherCalChoice.CheckedChanged += new System.EventHandler(this.radioButton_OtherCalChoice_CheckedChanged);
            // 
            // textBox_OtherCalName
            // 
            this.textBox_OtherCalName.Enabled = false;
            this.textBox_OtherCalName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.textBox_OtherCalName.Location = new System.Drawing.Point(298, 26);
            this.textBox_OtherCalName.Name = "textBox_OtherCalName";
            this.textBox_OtherCalName.Size = new System.Drawing.Size(143, 21);
            this.textBox_OtherCalName.TabIndex = 3;
            this.textBox_OtherCalName.Text = "Lotus.Notes";
            this.textBox_OtherCalName.LostFocus += new System.EventHandler(this.textBox_OtherCalName_LostFocus);
            // 
            // checkBox_ConnectUsingSSL
            // 
            this.checkBox_ConnectUsingSSL.AllowDrop = true;
            this.checkBox_ConnectUsingSSL.AutoSize = true;
            this.checkBox_ConnectUsingSSL.Checked = true;
            this.checkBox_ConnectUsingSSL.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ConnectUsingSSL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_ConnectUsingSSL.Location = new System.Drawing.Point(6, 53);
            this.checkBox_ConnectUsingSSL.Name = "checkBox_ConnectUsingSSL";
            this.checkBox_ConnectUsingSSL.Size = new System.Drawing.Size(240, 19);
            this.checkBox_ConnectUsingSSL.TabIndex = 4;
            this.checkBox_ConnectUsingSSL.Text = "Use SSL to connect to Google calendar";
            this.checkBox_ConnectUsingSSL.UseVisualStyleBackColor = true;
            // 
            // checkBox_CreateNotification
            // 
            this.checkBox_CreateNotification.AutoSize = true;
            this.checkBox_CreateNotification.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.checkBox_CreateNotification.Location = new System.Drawing.Point(6, 79);
            this.checkBox_CreateNotification.Name = "checkBox_CreateNotification";
            this.checkBox_CreateNotification.Size = new System.Drawing.Size(215, 19);
            this.checkBox_CreateNotification.TabIndex = 12;
            this.checkBox_CreateNotification.Text = "Create Notifications (30 mins prior)";
            this.checkBox_CreateNotification.UseVisualStyleBackColor = true;
            // 
            // tabPage_SchedulingSync
            // 
            this.tabPage_SchedulingSync.Controls.Add(this.groupBox_Sync);
            this.tabPage_SchedulingSync.Location = new System.Drawing.Point(4, 22);
            this.tabPage_SchedulingSync.Name = "tabPage_SchedulingSync";
            this.tabPage_SchedulingSync.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_SchedulingSync.Size = new System.Drawing.Size(572, 236);
            this.tabPage_SchedulingSync.TabIndex = 2;
            this.tabPage_SchedulingSync.Text = "Scheduling";
            this.tabPage_SchedulingSync.UseVisualStyleBackColor = true;
            // 
            // groupBox_Sync
            // 
            this.groupBox_Sync.Controls.Add(this.checkBox_ScheduleSync);
            this.groupBox_Sync.Controls.Add(this.label_Minutes);
            this.groupBox_Sync.Controls.Add(this.checkBox_SyncOnStartup);
            this.groupBox_Sync.Controls.Add(this.textBox_ScheduleSync);
            this.groupBox_Sync.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox_Sync.Location = new System.Drawing.Point(15, 17);
            this.groupBox_Sync.Name = "groupBox_Sync";
            this.groupBox_Sync.Size = new System.Drawing.Size(539, 83);
            this.groupBox_Sync.TabIndex = 5;
            this.groupBox_Sync.TabStop = false;
            this.groupBox_Sync.Text = "Synchronizing";
            // 
            // checkBox_ScheduleSync
            // 
            this.checkBox_ScheduleSync.AutoSize = true;
            this.checkBox_ScheduleSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.checkBox_ScheduleSync.Location = new System.Drawing.Point(6, 51);
            this.checkBox_ScheduleSync.Name = "checkBox_ScheduleSync";
            this.checkBox_ScheduleSync.Size = new System.Drawing.Size(107, 19);
            this.checkBox_ScheduleSync.TabIndex = 2;
            this.checkBox_ScheduleSync.Text = "Schedule Sync";
            this.checkBox_ScheduleSync.UseVisualStyleBackColor = true;
            this.checkBox_ScheduleSync.CheckedChanged += new System.EventHandler(this.checkBox_ScheduleSync_CheckedChanged);
            // 
            // label_Minutes
            // 
            this.label_Minutes.AutoSize = true;
            this.label_Minutes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.label_Minutes.Location = new System.Drawing.Point(219, 52);
            this.label_Minutes.Name = "label_Minutes";
            this.label_Minutes.Size = new System.Drawing.Size(51, 15);
            this.label_Minutes.TabIndex = 4;
            this.label_Minutes.Text = "minutes";
            // 
            // checkBox_SyncOnStartup
            // 
            this.checkBox_SyncOnStartup.AutoSize = true;
            this.checkBox_SyncOnStartup.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.checkBox_SyncOnStartup.Location = new System.Drawing.Point(6, 25);
            this.checkBox_SyncOnStartup.Name = "checkBox_SyncOnStartup";
            this.checkBox_SyncOnStartup.Size = new System.Drawing.Size(111, 19);
            this.checkBox_SyncOnStartup.TabIndex = 1;
            this.checkBox_SyncOnStartup.Text = "Sync on Startup";
            this.checkBox_SyncOnStartup.UseVisualStyleBackColor = true;
            // 
            // textBox_ScheduleSync
            // 
            this.textBox_ScheduleSync.Enabled = false;
            this.textBox_ScheduleSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.textBox_ScheduleSync.Location = new System.Drawing.Point(182, 49);
            this.textBox_ScheduleSync.Name = "textBox_ScheduleSync";
            this.textBox_ScheduleSync.Size = new System.Drawing.Size(31, 21);
            this.textBox_ScheduleSync.TabIndex = 3;
            this.textBox_ScheduleSync.Text = "30";
            this.textBox_ScheduleSync.LostFocus += new System.EventHandler(this.textBox_ScheduleSync_LostFocus);
            // 
            // tabPage_Debug
            // 
            this.tabPage_Debug.AutoScroll = true;
            this.tabPage_Debug.Controls.Add(this.checkBox_MessageBoxDebug);
            this.tabPage_Debug.Controls.Add(this.button_ClearDebug);
            this.tabPage_Debug.Controls.Add(this.textBox_Debug);
            this.tabPage_Debug.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Debug.Name = "tabPage_Debug";
            this.tabPage_Debug.Size = new System.Drawing.Size(572, 236);
            this.tabPage_Debug.TabIndex = 3;
            this.tabPage_Debug.Text = "Debug";
            this.tabPage_Debug.UseVisualStyleBackColor = true;
            // 
            // checkBox_MessageBoxDebug
            // 
            this.checkBox_MessageBoxDebug.AutoSize = true;
            this.checkBox_MessageBoxDebug.Location = new System.Drawing.Point(4, 211);
            this.checkBox_MessageBoxDebug.Name = "checkBox_MessageBoxDebug";
            this.checkBox_MessageBoxDebug.Size = new System.Drawing.Size(110, 17);
            this.checkBox_MessageBoxDebug.TabIndex = 2;
            this.checkBox_MessageBoxDebug.Text = "Prompt Messages";
            this.checkBox_MessageBoxDebug.UseVisualStyleBackColor = true;
            // 
            // button_ClearDebug
            // 
            this.button_ClearDebug.Location = new System.Drawing.Point(492, 210);
            this.button_ClearDebug.Name = "button_ClearDebug";
            this.button_ClearDebug.Size = new System.Drawing.Size(75, 23);
            this.button_ClearDebug.TabIndex = 1;
            this.button_ClearDebug.Text = "Clear";
            this.button_ClearDebug.UseVisualStyleBackColor = true;
            this.button_ClearDebug.Click += new System.EventHandler(this.button_ClearDebug_Click);
            // 
            // textBox_Debug
            // 
            this.textBox_Debug.AcceptsReturn = true;
            this.textBox_Debug.AcceptsTab = true;
            this.textBox_Debug.Location = new System.Drawing.Point(3, 13);
            this.textBox_Debug.Multiline = true;
            this.textBox_Debug.Name = "textBox_Debug";
            this.textBox_Debug.ReadOnly = true;
            this.textBox_Debug.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox_Debug.Size = new System.Drawing.Size(566, 191);
            this.textBox_Debug.TabIndex = 0;
            this.textBox_Debug.TextChanged += new System.EventHandler(this.textBox_Debug_TextChanged);
            // 
            // label_Author
            // 
            this.label_Author.AutoSize = true;
            this.label_Author.Location = new System.Drawing.Point(333, 9);
            this.label_Author.Name = "label_Author";
            this.label_Author.Size = new System.Drawing.Size(49, 13);
            this.label_Author.TabIndex = 1;
            this.label_Author.Text = "Website:";
            // 
            // label_Version
            // 
            this.label_Version.AutoSize = true;
            this.label_Version.Location = new System.Drawing.Point(333, 31);
            this.label_Version.Name = "label_Version";
            this.label_Version.Size = new System.Drawing.Size(45, 13);
            this.label_Version.TabIndex = 2;
            this.label_Version.Text = "Version:";
            // 
            // label_Date
            // 
            this.label_Date.AutoSize = true;
            this.label_Date.Location = new System.Drawing.Point(333, 53);
            this.label_Date.Name = "label_Date";
            this.label_Date.Size = new System.Drawing.Size(33, 13);
            this.label_Date.TabIndex = 3;
            this.label_Date.Text = "Date:";
            // 
            // label_VersionValue
            // 
            this.label_VersionValue.AutoSize = true;
            this.label_VersionValue.Location = new System.Drawing.Point(403, 31);
            this.label_VersionValue.Name = "label_VersionValue";
            this.label_VersionValue.Size = new System.Drawing.Size(0, 13);
            this.label_VersionValue.TabIndex = 5;
            // 
            // label_DateValue
            // 
            this.label_DateValue.AutoSize = true;
            this.label_DateValue.Location = new System.Drawing.Point(403, 53);
            this.label_DateValue.Name = "label_DateValue";
            this.label_DateValue.Size = new System.Drawing.Size(0, 13);
            this.label_DateValue.TabIndex = 6;
            // 
            // label_About
            // 
            this.label_About.AutoSize = true;
            this.label_About.Location = new System.Drawing.Point(333, 76);
            this.label_About.Name = "label_About";
            this.label_About.Size = new System.Drawing.Size(38, 13);
            this.label_About.TabIndex = 8;
            this.label_About.Text = "About:";
            // 
            // textBox_AboutValue
            // 
            this.textBox_AboutValue.BackColor = System.Drawing.SystemColors.Control;
            this.textBox_AboutValue.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_AboutValue.ForeColor = System.Drawing.SystemColors.ControlText;
            this.textBox_AboutValue.Location = new System.Drawing.Point(403, 77);
            this.textBox_AboutValue.Multiline = true;
            this.textBox_AboutValue.Name = "textBox_AboutValue";
            this.textBox_AboutValue.Size = new System.Drawing.Size(180, 56);
            this.textBox_AboutValue.TabIndex = 9;
            // 
            // button_ManualSync
            // 
            this.button_ManualSync.Location = new System.Drawing.Point(515, 413);
            this.button_ManualSync.Name = "button_ManualSync";
            this.button_ManualSync.Size = new System.Drawing.Size(75, 23);
            this.button_ManualSync.TabIndex = 10;
            this.button_ManualSync.Text = "Synchronize";
            this.button_ManualSync.UseVisualStyleBackColor = true;
            this.button_ManualSync.Click += new System.EventHandler(this.button_ManualSync_Click);
            // 
            // button_SaveSettings
            // 
            this.button_SaveSettings.Location = new System.Drawing.Point(434, 413);
            this.button_SaveSettings.Name = "button_SaveSettings";
            this.button_SaveSettings.Size = new System.Drawing.Size(75, 23);
            this.button_SaveSettings.TabIndex = 11;
            this.button_SaveSettings.Text = "Save";
            this.button_SaveSettings.UseVisualStyleBackColor = true;
            this.button_SaveSettings.Click += new System.EventHandler(this.button_SaveSettings_Click);
            // 
            // panel_StatusPanel
            // 
            this.panel_StatusPanel.Controls.Add(this.progressBar_SyncProgress);
            this.panel_StatusPanel.Controls.Add(this.label_Status);
            this.panel_StatusPanel.Location = new System.Drawing.Point(12, 413);
            this.panel_StatusPanel.Name = "panel_StatusPanel";
            this.panel_StatusPanel.Size = new System.Drawing.Size(304, 23);
            this.panel_StatusPanel.TabIndex = 12;
            // 
            // progressBar_SyncProgress
            // 
            this.progressBar_SyncProgress.Location = new System.Drawing.Point(74, 4);
            this.progressBar_SyncProgress.Name = "progressBar_SyncProgress";
            this.progressBar_SyncProgress.Size = new System.Drawing.Size(227, 16);
            this.progressBar_SyncProgress.Step = 5;
            this.progressBar_SyncProgress.TabIndex = 1;
            // 
            // label_Status
            // 
            this.label_Status.AutoSize = true;
            this.label_Status.Location = new System.Drawing.Point(3, 5);
            this.label_Status.Name = "label_Status";
            this.label_Status.Size = new System.Drawing.Size(64, 13);
            this.label_Status.TabIndex = 0;
            this.label_Status.Text = "Sync Status";
            // 
            // pictureBox_Logo
            // 
            this.pictureBox_Logo.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox_Logo.Image")));
            this.pictureBox_Logo.ImageLocation = "";
            this.pictureBox_Logo.Location = new System.Drawing.Point(12, 9);
            this.pictureBox_Logo.Name = "pictureBox_Logo";
            this.pictureBox_Logo.Size = new System.Drawing.Size(300, 124);
            this.pictureBox_Logo.TabIndex = 7;
            this.pictureBox_Logo.TabStop = false;
            // 
            // backgroundWorker_SingleSync
            // 
            this.backgroundWorker_SingleSync.WorkerReportsProgress = true;
            this.backgroundWorker_SingleSync.DoWork += new System.ComponentModel.DoWorkEventHandler(this.DoTheWork);
            this.backgroundWorker_SingleSync.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.DoTheWork_Completed);
            this.backgroundWorker_SingleSync.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.DoTheWork_Progress);
            // 
            // notifyIcon_Tray
            // 
            this.notifyIcon_Tray.BalloonTipText = "iNotes to Google";
            this.notifyIcon_Tray.BalloonTipTitle = "iNotes to Google";
            this.notifyIcon_Tray.ContextMenuStrip = this.contextMenu_Tray;
            this.notifyIcon_Tray.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon_Tray.Icon")));
            this.notifyIcon_Tray.Text = "iNotes to Google";
            this.notifyIcon_Tray.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_Tray_MouseDoubleClick);
            // 
            // contextMenu_Tray
            // 
            this.contextMenu_Tray.BackColor = System.Drawing.SystemColors.Menu;
            this.contextMenu_Tray.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.contextMenu_Tray.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem_Open,
            this.toolStripMenuItem_Sync,
            this.toolStripSeparator1,
            this.toolStripMenuItem_Close});
            this.contextMenu_Tray.Name = "contextMenu_Tray";
            this.contextMenu_Tray.ShowImageMargin = false;
            this.contextMenu_Tray.Size = new System.Drawing.Size(86, 76);
            // 
            // toolStripMenuItem_Open
            // 
            this.toolStripMenuItem_Open.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripMenuItem_Open.Name = "toolStripMenuItem_Open";
            this.toolStripMenuItem_Open.Size = new System.Drawing.Size(85, 22);
            this.toolStripMenuItem_Open.Text = "Open";
            this.toolStripMenuItem_Open.Click += new System.EventHandler(this.toolStripMenuItem_Open_Click);
            // 
            // toolStripMenuItem_Sync
            // 
            this.toolStripMenuItem_Sync.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripMenuItem_Sync.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold);
            this.toolStripMenuItem_Sync.Name = "toolStripMenuItem_Sync";
            this.toolStripMenuItem_Sync.Size = new System.Drawing.Size(85, 22);
            this.toolStripMenuItem_Sync.Text = "Sync";
            this.toolStripMenuItem_Sync.Click += new System.EventHandler(this.toolStripMenuItem_Sync_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(82, 6);
            // 
            // toolStripMenuItem_Close
            // 
            this.toolStripMenuItem_Close.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripMenuItem_Close.Name = "toolStripMenuItem_Close";
            this.toolStripMenuItem_Close.Size = new System.Drawing.Size(85, 22);
            this.toolStripMenuItem_Close.Text = "Close";
            this.toolStripMenuItem_Close.Click += new System.EventHandler(this.toolStripMenuItem_Close_Click);
            // 
            // linkLabel_url
            // 
            this.linkLabel_url.AutoSize = true;
            this.linkLabel_url.Location = new System.Drawing.Point(403, 9);
            this.linkLabel_url.Name = "linkLabel_url";
            this.linkLabel_url.Size = new System.Drawing.Size(130, 13);
            this.linkLabel_url.TabIndex = 13;
            this.linkLabel_url.TabStop = true;
            this.linkLabel_url.Text = "Sourceforge / iNotesSync";
            this.linkLabel_url.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel_url_LinkClicked);
            // 
            // NotesToGoogleForm
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(602, 446);
            this.Controls.Add(this.linkLabel_url);
            this.Controls.Add(this.panel_StatusPanel);
            this.Controls.Add(this.button_SaveSettings);
            this.Controls.Add(this.button_ManualSync);
            this.Controls.Add(this.textBox_AboutValue);
            this.Controls.Add(this.label_About);
            this.Controls.Add(this.pictureBox_Logo);
            this.Controls.Add(this.label_DateValue);
            this.Controls.Add(this.label_VersionValue);
            this.Controls.Add(this.label_Date);
            this.Controls.Add(this.label_Version);
            this.Controls.Add(this.label_Author);
            this.Controls.Add(this.tabControl_MainControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NotesToGoogleForm";
            this.Text = "Notes To Google Calender Sync";
            this.TransparencyKey = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.Load += new System.EventHandler(this.NotesToGoogleForm_Load);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.FormClosingEventCancel_Closing);
            this.tabControl_MainControl.ResumeLayout(false);
            this.tabPage_Connect.ResumeLayout(false);
            this.groupBox_GoogleConnection.ResumeLayout(false);
            this.groupBox_GoogleConnection.PerformLayout();
            this.groupBox_LotusConnection.ResumeLayout(false);
            this.groupBox_LotusConnection.PerformLayout();
            this.tabPage_Preferences.ResumeLayout(false);
            this.groupBox_language.ResumeLayout(false);
            this.groupBox_language.PerformLayout();
            this.groupBox_Window.ResumeLayout(false);
            this.groupBox_Window.PerformLayout();
            this.groupBox_LotusNotes.ResumeLayout(false);
            this.groupBox_LotusNotes.PerformLayout();
            this.groupBox_GoogleCalendar.ResumeLayout(false);
            this.groupBox_GoogleCalendar.PerformLayout();
            this.tabPage_SchedulingSync.ResumeLayout(false);
            this.groupBox_Sync.ResumeLayout(false);
            this.groupBox_Sync.PerformLayout();
            this.tabPage_Debug.ResumeLayout(false);
            this.tabPage_Debug.PerformLayout();
            this.panel_StatusPanel.ResumeLayout(false);
            this.panel_StatusPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Logo)).EndInit();
            this.contextMenu_Tray.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        ///
        /// GUI Developer class variables
        ///
        private int newDaysAhead;
        private TabControl tabControl_MainControl;
        private TabPage tabPage_Connect;
        private PictureBox pictureBox_Logo;
        private TextBox textBox_GooglePassword;
        private TextBox textBox_GoogleLogin;
        private TabPage tabPage_Preferences;
        private Label label_Author;
        private Label label_Version;
        private Label label_Date;
        private Label label_VersionValue;
        private Label label_DateValue;
        private Label label_About;
        private TextBox textBox_AboutValue;
        private Button button_ManualSync;
        private Button button_SaveSettings;
        private Panel panel_StatusPanel;
        private ProgressBar progressBar_SyncProgress;
        private Label label_Status;
        private Label label_NotesWebmailURL;
        private Label label_NotesPassword;
        private Label label_NotesLogin;
        private Label label_GooglePassword;
        private Label label_GoogleLogin;
        private TextBox textBox_NotesLogin;
        private TextBox textBox_WebmailURL;
        private RadioButton radioButton_MainCalChoice;
        private TextBox textBox_OtherCalName;
        private RadioButton radioButton_OtherCalChoice;
        private CheckBox checkBox_ConnectUsingSSL;
        private CheckBox checkBox_NotesServerAuth;
        private TabPage tabPage_SchedulingSync;
        private CheckBox checkBox_SyncOnStartup;
        private Label label_Minutes;
        private TextBox textBox_ScheduleSync;
        private CheckBox checkBox_ScheduleSync;
        private TextBox textBox_NotesPassword;
        private CheckBox checkBox_MinimizeToTray;
        private TabPage tabPage_Debug;
        private TextBox textBox_Debug;
        private Button button_ClearDebug;
        private TextBox textBox_CustomDaysAhead;
        private CheckBox checkBox_CustomDaysAhead;
        private Button button_CreateService;
        private SyncPreferences config;
        private CheckBox checkBox_MessageBoxDebug;
        private CheckBox checkBox_CreateNotification;
        private BackgroundWorker backgroundWorker_SingleSync;
        private NotifyIcon notifyIcon_Tray;
        private IContainer components;
        private ContextMenuStrip contextMenu_Tray;
        private ToolStripMenuItem toolStripMenuItem_Open;
        private ToolStripMenuItem toolStripMenuItem_Sync;
        private ToolStripMenuItem toolStripMenuItem_Close;
        private ToolStripSeparator toolStripSeparator1;
        private CheckBox checkBox_StartMinimized;
        private LinkLabel linkLabel_url;
        private RadioButton radioButton_French;
        private RadioButton radioButton_English;
        private GroupBox groupBox_GoogleCalendar;
        private GroupBox groupBox_LotusNotes;
        private GroupBox groupBox_Window;
        private GroupBox groupBox_language;
        private GroupBox groupBox_LotusConnection;
        private GroupBox groupBox_GoogleConnection;
        private GroupBox groupBox_Sync;
        private BackgroundWorker backgroundWorker_Schedule;
        private System.Windows.Forms.Timer timer_Schedule;
        

        /// <summary>
        /// Default days ahead value, constant
        /// </summary>
        public const int DEFAULTDAYSAHEAD = 14;

        /// <summary>
        /// Event handler called when the radio button for Other Calendar event is changed.
        /// This event will set the Other calendar name text box enabled or not.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton_OtherCalChoice_CheckedChanged(object sender, EventArgs e)
        {
            textBox_OtherCalName.Enabled = radioButton_OtherCalChoice.Checked;
        }

        /// <summary>
        /// Event method called when notifyIcon is double clicked
        /// </summary>
        /// <param name="sender">Sender object</param>
        /// <param name="e">MouseEvent arguments passed</param>
        private void notifyIcon_Tray_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            notifyIcon_Tray.Visible = false;
        }

        /// <summary>
        /// Event handler used to clear the Debug text box when pressed
        /// </summary>
        /// <param name="sender">Object that sent the Event</param>
        /// <param name="e">Event arguments passed by Event</param>
        private void button_ClearDebug_Click(object sender, EventArgs e)
        {
            textBox_Debug.Clear();
        }

        /// <summary>
        /// Event handler called when focus is taken off of Custom Days text box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_CustomDaysAhead_LostFocus(object sender, EventArgs e)
        {
            try
            {
                newDaysAhead = int.Parse(textBox_CustomDaysAhead.Text);
            }
            catch (FormatException ex)
            {
                PrintStringToDebug("Days Ahead: Value is not an appropriate number. Reverting to old value.");
                textBox_CustomDaysAhead.Text = (config.GetPreference("DaysAhead") != "") ? config.GetPreference("DaysAhead") : "14";              
            }
        }

        /// <summary>
        /// Event handler method called when the user changes the value of Custom Days Ahead Check
        /// </summary>
        /// <param name="sender">Object sending the request</param>
        /// <param name="e">Event Args being passed</param>
        private void checkBox_CustomDaysAhead_CheckedChanged(object sender, EventArgs e)
        {
            textBox_CustomDaysAhead.Enabled = checkBox_CustomDaysAhead.Checked;
        }

        /// <summary>
        /// Event handler called when the radio button for Main Calendar Choice is changed.
        /// This event will set the Other calendar name text box enabled or not.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton_MainCalChoice_CheckedChanged(object sender, EventArgs e)
        {
            textBox_OtherCalName.Enabled = !radioButton_MainCalChoice.Checked;
        }

        /// <summary>
        /// Event handler for when other cal loses focus, we need to check to see that the calendar has an
        /// appropriate name
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_OtherCalName_LostFocus(object sender, EventArgs e)
        {
            if (textBox_OtherCalName.Text == "")
            {
                textBox_OtherCalName.Text = "Lotus.Notes";
            }
        }

        /// <summary>
        /// Event handler when the schedule sync is changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox_ScheduleSync_CheckedChanged(object sender, EventArgs e)
        {
            textBox_ScheduleSync.Enabled = checkBox_ScheduleSync.Checked;
            PrintStringToDebug("Please restart to change Schedule settings.");
            //SetupAndStartScheduleSync();
        }

        /// <summary>
        /// Method used to change the language to French
        /// </summary>
        /// <param name="sender">Sender of the request</param>
        /// <param name="e">Event args passed from sender</param>
        private void radioButton_French_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_French.Checked)
            {
                Thread.CurrentThread.CurrentUICulture = m_FrenchCulture;
                UpdateUI();
            }
        }

        /// <summary>
        /// Method used to change the language to English
        /// </summary>
        /// <param name="sender">Sender of the request</param>
        /// <param name="e">Event args passed from sender</param>
        private void radioButton_English_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_English.Checked)
            {
                Thread.CurrentThread.CurrentUICulture = m_EnglishCulture;
                UpdateUI();
            }
        }

        /// <summary>
        /// Event handler used when Save button is clicked
        /// </summary>
        /// <param name="sender">Object that initiated the Event</param>
        /// <param name="e">Event arguments passed by Event</param>
        private void button_SaveSettings_Click(object sender, EventArgs e)
        {
            // Save Connection Variables
            config.SetPreference("WebmailURL", textBox_WebmailURL.Text);
            config.SetPreference("NotesLogin", textBox_NotesLogin.Text);
            config.SetPreference("NotesPassword", textBox_NotesPassword.Text);
            config.SetPreference("GoogleLogin", textBox_GoogleLogin.Text);
            config.SetPreference("GooglePassword", textBox_GooglePassword.Text);

            // Save Preference Variables
            config.SetPreference("ConnectUsingSSL", checkBox_ConnectUsingSSL.Checked.ToString());
            config.SetPreference("NotesServerAuth", checkBox_NotesServerAuth.Checked.ToString());
            config.SetPreference("MinimizeToTray", checkBox_MinimizeToTray.Checked.ToString());
            config.SetPreference("MainCalChoice", radioButton_MainCalChoice.Checked.ToString());
            config.SetPreference("OtherCalChoice", radioButton_OtherCalChoice.Checked.ToString());
            config.SetPreference("OtherCalName", textBox_OtherCalName.Text);
            config.SetPreference("CustomDaysAhead", checkBox_CustomDaysAhead.Checked.ToString());
            config.SetPreference("NotificationsOn", checkBox_CreateNotification.Checked.ToString());
            config.SetPreference("DaysAhead", textBox_CustomDaysAhead.Text);
            config.SetPreference("StartMinimized", checkBox_StartMinimized.Checked.ToString());

            // Save Schedule Variables
            config.SetPreference("SyncOnStartup", checkBox_SyncOnStartup.Checked.ToString());
            config.SetPreference("ScheduleSync", checkBox_ScheduleSync.Checked.ToString());
            config.SetPreference("ScheduleTime", textBox_ScheduleSync.Text);

            // Save Config File
            config.SavePreferences();

            // Confirm that file was saved
            PrintStringToDebug("Configuration was saved.");
        }

        /// <summary>
        /// Event handler called when website link is clicked
        /// </summary>
        /// <param name="sender">Sender of the event</param>
        /// <param name="e">Parameters passed</param>
        private void linkLabel_url_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("iexplore.exe", "https://sourceforge.net/projects/inotessync/");
        }

        /// <summary>
        /// Method used to check size and clear if needed the debug screen
        /// </summary>
        /// <param name="sender">sender of the request</param>
        /// <param name="e">event argumetns passed</param>
        private void textBox_Debug_TextChanged(object sender, EventArgs e)
        {   // Control the size of the textarea
            if (textBox_Debug.Text.Length > 5000)
            {
                textBox_Debug.Text = "";
            }
        }

        /// <summary>
        /// Method called when text box is changed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox_ScheduleSync_LostFocus(object sender, EventArgs e)
        {
            try
            {
                int testvalue = int.Parse(textBox_ScheduleSync.Text);
            }
            catch (FormatException ex)
            {
                PrintStringToDebug("Schedule Time: Value is not an appropriate number. Reverting to last saved value.");
                textBox_ScheduleSync.Text = (config.GetPreference("ScheduleTime") != "") ? config.GetPreference("ScheduleTime") : "30";
            }
            PrintStringToDebug("You must Save and Restart the application to sync with \r\nthe new interval.");
        }

        /// <summary>
        /// Private method used to update the Progress bar while syncing.  This should be called by a thread
        /// safe delegate in the DoTheWork thread method.
        /// </summary>
        /// <param name="value">value to increase progress bar</param>
        private void UpdateSyncProgressBar(int value)
        {
            this.progressBar_SyncProgress.Value = value;
        }

        /// <summary>
        /// Event handler called when the TrayMenu option Closed is selected;  this will submit a close
        /// request to the Form
        /// </summary>
        /// <param name="sender">sender of the request</param>
        /// <param name="e">events passed through the object</param>
        private void toolStripMenuItem_Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Event handler called when the TrayMenu option Sync is selected; this will call a button_manualsync
        /// event.
        /// </summary>
        /// <param name="sender">sender of the event</param>
        /// <param name="e">arguments passed through the object</param>
        private void toolStripMenuItem_Sync_Click(object sender, EventArgs e)
        {
            button_ManualSync_Click((object)toolStripMenuItem_Sync, e);
        }

        /// <summary>
        /// Event handler called when the TrayMenu option Open is selected; this will call TrayIcon
        /// event hanlder.
        /// </summary>
        /// <param name="sender">sender of the event</param>
        /// <param name="e">arguments from the sender</param>
        private void toolStripMenuItem_Open_Click(object sender, EventArgs e)
        {
            notifyIcon_Tray_MouseDoubleClick((object)toolStripMenuItem_Sync, (MouseEventArgs) null);
        }

        /// <summary>
        /// Method called when Manual Sync buttong has been pushed
        /// </summary>
        /// <param name="sender">Object creating the event</param>
        /// <param name="e">Event arguments</param>
        private void button_ManualSync_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker_SingleSync.IsBusy)
            {   // Check to see if the background worker is already started, this will ensure two syncs are never
                // done at the same time.
                // Disable the sync button, so we don't do this twice.
                button_ManualSync.Enabled = false;

                // Start the background worker.
                backgroundWorker_SingleSync.RunWorkerAsync();
            }
        }
        
        /// <summary>
        /// Event method called by the schedule timer.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_ScheduleSync_Tick(Object sender, EventArgs e)
        {
            if (!backgroundWorker_SingleSync.IsBusy)
            {   // Check to see if the background worker is already started, this will ensure two syncs are never
                // done at the same time.
                // Disable the sync button, so we don't do this twice.
                button_ManualSync.Enabled = false;

                // Start the background worker.
                backgroundWorker_SingleSync.RunWorkerAsync();
            }
        }

        #region Thread_Methods
        //////////////////////////// THREAD SECTION ////////////////////////////////////////////////////

        /// <summary>
        /// Completed event handler for background workers
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DoTheWork_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {   // EventArg e contained an error
                PrintStringToDebug("Error during sync: " + e.Error.Message);
            }
            else
            {
                // If the Sync completed successfully; no errors within the EventArgs
                PrintStringToDebug("Sync completed successfully; synced " + e.Result + " calendar entries.");
            }

            // Enable the sync button so we can redo the sync
            button_ManualSync.Enabled = true;
            progressBar_SyncProgress.Value = 0;
        }

        /// <summary>
        /// Progress Changed event handler for background workers
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DoTheWork_Progress(object sender, ProgressChangedEventArgs e)
        {   // Check the progress changed events for ProgressPercentage
            // update the progress bar with the new value

            progressBar_SyncProgress.Value = e.ProgressPercentage;

        }

        /// <summary>
        /// DoWork method for Background workers, method must not attempt to touch the UI, we use Progress
        /// Changed for that.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        [System.Diagnostics.DebuggerNonUserCodeAttribute()] 
        private void DoTheWork(object sender, DoWorkEventArgs e)
        {
            // Get the background worker as an object
            BackgroundWorker bWorker = sender as BackgroundWorker;

            // create a temp calendar, for testing purposes
            NotesToGoogleSync sync;
            int daysAheadToCheck = (checkBox_CustomDaysAhead.Checked) ? int.Parse(textBox_CustomDaysAhead.Text) : DEFAULTDAYSAHEAD;
            string defaultCalendar = (radioButton_MainCalChoice.Checked) ? textBox_GoogleLogin.Text : textBox_OtherCalName.Text;           

            // Create the Google service
            GoogleServiceConnect googleService = new GoogleServiceConnect(
                 textBox_GoogleLogin.Text, textBox_GooglePassword.Text, 
                 checkBox_ConnectUsingSSL.Checked, checkBox_CreateNotification.Checked);
            googleService.Calendar = defaultCalendar;
            googleService.DaysAhead = daysAheadToCheck;

            bWorker.ReportProgress(5);

            // Create the Lotus Notes service
            NotesServiceConnect notesService = new NotesServiceConnect(
                textBox_NotesLogin.Text, textBox_NotesPassword.Text, textBox_WebmailURL.Text, daysAheadToCheck);

            bWorker.ReportProgress(10);

            // Do some checking of services and connectivity
            // If successful do the sync
            if (googleService == null || !googleService.TestConnection())
            {
                e.Result = NotesToGoogleSync.SYNCNULLGSERVICE;
                throw new Exception("Google service not initialized; check login information");               
            }
            else if (notesService == null || !notesService.TestConnection())
            {
                e.Result = NotesToGoogleSync.SYNCNULLNSERVICE;
                throw new Exception("Notes service not initialized; check login information");
            }
            else
            {
                // Create the notes sync object and start a sync
                // This is the bulk of the work; will have to add a delegate within the NotesToGoogleSync
                // Class that can be set to invoke the status bar
                sync = new NotesToGoogleSync(googleService, notesService);
                bWorker.ReportProgress(15);

                e.Result = sync.Sync(bWorker, e);
                notesService.CloseConnection();
            }
        }
        #endregion
     
        #region User_Variables
        ///
        /// Setup user variables for language localization
        ///
        private ResourceManager m_ResourceManager = new ResourceManager("NotesToGoogle.Localization",
                                    System.Reflection.Assembly.GetExecutingAssembly());
        private CultureInfo m_EnglishCulture = new CultureInfo("en-CA");
        private CultureInfo m_FrenchCulture = new CultureInfo("fr-FR");

        #endregion

        // End of the NotesToGoogleAppClass
    }
}
