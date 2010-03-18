using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
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
            config = new SyncPreferences();
            InitializeComponent();
            m_previousWindowState = (this.WindowState == FormWindowState.Minimized ? FormWindowState.Normal : this.WindowState);
        }

        /// <summary>
        /// Override method from parent.OnResize; method will move application to tray
        /// </summary>
        /// <param name="e">Event args passed to on resize</param>
        protected override void OnResize(EventArgs e)
        {   // Call the parent OnResize
            base.OnResize(e);

            if (checkBox_MinimizeToTray.Checked)
            {
                // We need to keep track of whether you're minimizing from a normal or maximized window
                if (this.WindowState != FormWindowState.Minimized)
                {
                    m_previousWindowState = this.WindowState;
                }
                notifyIcon_Tray.Visible = (this.WindowState == FormWindowState.Minimized);
                this.Visible = !notifyIcon_Tray.Visible;
            }
        }

        /// <summary>
        /// Method called when the main form is loaded
        /// </summary>
        /// <param name="sender">Object taht sent the Load event</param>
        /// <param name="e">Event Arguments passed</param>
        private void NotesToGoogleForm_Load(object sender, EventArgs e)
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

                // Save Schedule Variables
                checkBox_SyncOnStartup.Checked = (config.GetPreference("SyncOnStartup") != "") ? Convert.ToBoolean(config.GetPreference("SyncOnStartup")) : false;
                checkBox_ScheduleSync.Checked = (config.GetPreference("ScheduleSync") != "") ? Convert.ToBoolean(config.GetPreference("ScheduleSync")) : false;
            }
            else
            {   // If there were errors, we need to post an error
                PrintStringToDebug("Config Preferences did not load properly");
            }

            // Sync on Startup if selected
            if (checkBox_SyncOnStartup.Checked)
            {   // If the user has checked to Sync on Startup, we simulate the Button click
                button_ManualSync_Click((object)null, (EventArgs)null);
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
            this.textBox_GooglePassword = new System.Windows.Forms.TextBox();
            this.textBox_GoogleLogin = new System.Windows.Forms.TextBox();
            this.textBox_NotesPassword = new System.Windows.Forms.TextBox();
            this.textBox_NotesLogin = new System.Windows.Forms.TextBox();
            this.textBox_WebmailURL = new System.Windows.Forms.TextBox();
            this.label_GooglePassword = new System.Windows.Forms.Label();
            this.label_GoogleLogin = new System.Windows.Forms.Label();
            this.label_NotesPassword = new System.Windows.Forms.Label();
            this.label_NotesLogin = new System.Windows.Forms.Label();
            this.label_NotesWebmailURL = new System.Windows.Forms.Label();
            this.label_GoogleSection = new System.Windows.Forms.Label();
            this.label_NotesSection = new System.Windows.Forms.Label();
            this.tabPage_Preferences = new System.Windows.Forms.TabPage();
            this.checkBox_CreateNotification = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox_CustomDaysAhead = new System.Windows.Forms.TextBox();
            this.checkBox_CustomDaysAhead = new System.Windows.Forms.CheckBox();
            this.checkBox_MinimizeToTray = new System.Windows.Forms.CheckBox();
            this.label_WindowSection = new System.Windows.Forms.Label();
            this.checkBox_NotesServerAuth = new System.Windows.Forms.CheckBox();
            this.label_NotesConnectionSection = new System.Windows.Forms.Label();
            this.checkBox_ConnectUsingSSL = new System.Windows.Forms.CheckBox();
            this.textBox_OtherCalName = new System.Windows.Forms.TextBox();
            this.radioButton_OtherCalChoice = new System.Windows.Forms.RadioButton();
            this.radioButton_MainCalChoice = new System.Windows.Forms.RadioButton();
            this.label_CalendarSection = new System.Windows.Forms.Label();
            this.tabPage_SchedulingSync = new System.Windows.Forms.TabPage();
            this.label_Minutes = new System.Windows.Forms.Label();
            this.textBox_ScheduleSync = new System.Windows.Forms.TextBox();
            this.checkBox_ScheduleSync = new System.Windows.Forms.CheckBox();
            this.checkBox_SyncOnStartup = new System.Windows.Forms.CheckBox();
            this.label_SchedulingSection = new System.Windows.Forms.Label();
            this.tabPage_Debug = new System.Windows.Forms.TabPage();
            this.checkBox_MessageBoxDebug = new System.Windows.Forms.CheckBox();
            this.button_ClearDebug = new System.Windows.Forms.Button();
            this.textBox_Debug = new System.Windows.Forms.TextBox();
            this.label_Author = new System.Windows.Forms.Label();
            this.label_Version = new System.Windows.Forms.Label();
            this.label_Date = new System.Windows.Forms.Label();
            this.label_AuthorValue = new System.Windows.Forms.Label();
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
            this.tabControl_MainControl.SuspendLayout();
            this.tabPage_Connect.SuspendLayout();
            this.tabPage_Preferences.SuspendLayout();
            this.tabPage_SchedulingSync.SuspendLayout();
            this.tabPage_Debug.SuspendLayout();
            this.panel_StatusPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Logo)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl_MainControl
            // 
            this.tabControl_MainControl.Controls.Add(this.tabPage_Connect);
            this.tabControl_MainControl.Controls.Add(this.tabPage_Preferences);
            this.tabControl_MainControl.Controls.Add(this.tabPage_SchedulingSync);
            this.tabControl_MainControl.Controls.Add(this.tabPage_Debug);
            this.tabControl_MainControl.Location = new System.Drawing.Point(12, 139);
            this.tabControl_MainControl.Multiline = true;
            this.tabControl_MainControl.Name = "tabControl_MainControl";
            this.tabControl_MainControl.SelectedIndex = 0;
            this.tabControl_MainControl.Size = new System.Drawing.Size(580, 262);
            this.tabControl_MainControl.TabIndex = 0;
            // 
            // tabPage_Connect
            // 
            this.tabPage_Connect.Controls.Add(this.textBox_GooglePassword);
            this.tabPage_Connect.Controls.Add(this.textBox_GoogleLogin);
            this.tabPage_Connect.Controls.Add(this.textBox_NotesPassword);
            this.tabPage_Connect.Controls.Add(this.textBox_NotesLogin);
            this.tabPage_Connect.Controls.Add(this.textBox_WebmailURL);
            this.tabPage_Connect.Controls.Add(this.label_GooglePassword);
            this.tabPage_Connect.Controls.Add(this.label_GoogleLogin);
            this.tabPage_Connect.Controls.Add(this.label_NotesPassword);
            this.tabPage_Connect.Controls.Add(this.label_NotesLogin);
            this.tabPage_Connect.Controls.Add(this.label_NotesWebmailURL);
            this.tabPage_Connect.Controls.Add(this.label_GoogleSection);
            this.tabPage_Connect.Controls.Add(this.label_NotesSection);
            this.tabPage_Connect.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage_Connect.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Connect.Name = "tabPage_Connect";
            this.tabPage_Connect.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Connect.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.tabPage_Connect.Size = new System.Drawing.Size(572, 236);
            this.tabPage_Connect.TabIndex = 0;
            this.tabPage_Connect.Text = "Connection";
            this.tabPage_Connect.ToolTipText = "Set Connection Information";
            this.tabPage_Connect.UseVisualStyleBackColor = true;
            // 
            // textBox_GooglePassword
            // 
            this.textBox_GooglePassword.Location = new System.Drawing.Point(159, 186);
            this.textBox_GooglePassword.MaxLength = 20;
            this.textBox_GooglePassword.Name = "textBox_GooglePassword";
            this.textBox_GooglePassword.PasswordChar = '*';
            this.textBox_GooglePassword.Size = new System.Drawing.Size(177, 21);
            this.textBox_GooglePassword.TabIndex = 11;
            // 
            // textBox_GoogleLogin
            // 
            this.textBox_GoogleLogin.Location = new System.Drawing.Point(159, 159);
            this.textBox_GoogleLogin.MaxLength = 50;
            this.textBox_GoogleLogin.Name = "textBox_GoogleLogin";
            this.textBox_GoogleLogin.Size = new System.Drawing.Size(177, 21);
            this.textBox_GoogleLogin.TabIndex = 10;
            // 
            // textBox_NotesPassword
            // 
            this.textBox_NotesPassword.Location = new System.Drawing.Point(159, 89);
            this.textBox_NotesPassword.MaxLength = 20;
            this.textBox_NotesPassword.Name = "textBox_NotesPassword";
            this.textBox_NotesPassword.PasswordChar = '*';
            this.textBox_NotesPassword.Size = new System.Drawing.Size(177, 21);
            this.textBox_NotesPassword.TabIndex = 9;
            // 
            // textBox_NotesLogin
            // 
            this.textBox_NotesLogin.Location = new System.Drawing.Point(159, 63);
            this.textBox_NotesLogin.MaxLength = 20;
            this.textBox_NotesLogin.Name = "textBox_NotesLogin";
            this.textBox_NotesLogin.Size = new System.Drawing.Size(177, 21);
            this.textBox_NotesLogin.TabIndex = 8;
            // 
            // textBox_WebmailURL
            // 
            this.textBox_WebmailURL.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_WebmailURL.Location = new System.Drawing.Point(159, 37);
            this.textBox_WebmailURL.MaxLength = 100;
            this.textBox_WebmailURL.Name = "textBox_WebmailURL";
            this.textBox_WebmailURL.Size = new System.Drawing.Size(336, 21);
            this.textBox_WebmailURL.TabIndex = 7;
            // 
            // label_GooglePassword
            // 
            this.label_GooglePassword.AutoSize = true;
            this.label_GooglePassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_GooglePassword.Location = new System.Drawing.Point(15, 189);
            this.label_GooglePassword.Name = "label_GooglePassword";
            this.label_GooglePassword.Size = new System.Drawing.Size(61, 15);
            this.label_GooglePassword.TabIndex = 6;
            this.label_GooglePassword.Text = "Password";
            // 
            // label_GoogleLogin
            // 
            this.label_GoogleLogin.AutoSize = true;
            this.label_GoogleLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_GoogleLogin.Location = new System.Drawing.Point(15, 162);
            this.label_GoogleLogin.Name = "label_GoogleLogin";
            this.label_GoogleLogin.Size = new System.Drawing.Size(110, 15);
            this.label_GoogleLogin.TabIndex = 5;
            this.label_GoogleLogin.Text = "Username (email):";
            // 
            // label_NotesPassword
            // 
            this.label_NotesPassword.AutoSize = true;
            this.label_NotesPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_NotesPassword.Location = new System.Drawing.Point(15, 92);
            this.label_NotesPassword.Name = "label_NotesPassword";
            this.label_NotesPassword.Size = new System.Drawing.Size(64, 15);
            this.label_NotesPassword.TabIndex = 4;
            this.label_NotesPassword.Text = "Password:";
            // 
            // label_NotesLogin
            // 
            this.label_NotesLogin.AutoSize = true;
            this.label_NotesLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_NotesLogin.Location = new System.Drawing.Point(15, 66);
            this.label_NotesLogin.Name = "label_NotesLogin";
            this.label_NotesLogin.Size = new System.Drawing.Size(41, 15);
            this.label_NotesLogin.TabIndex = 3;
            this.label_NotesLogin.Text = "Login:";
            // 
            // label_NotesWebmailURL
            // 
            this.label_NotesWebmailURL.AutoSize = true;
            this.label_NotesWebmailURL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_NotesWebmailURL.Location = new System.Drawing.Point(15, 40);
            this.label_NotesWebmailURL.Name = "label_NotesWebmailURL";
            this.label_NotesWebmailURL.Size = new System.Drawing.Size(87, 15);
            this.label_NotesWebmailURL.TabIndex = 2;
            this.label_NotesWebmailURL.Text = "Webmail URL:";
            // 
            // label_GoogleSection
            // 
            this.label_GoogleSection.AutoSize = true;
            this.label_GoogleSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_GoogleSection.Location = new System.Drawing.Point(15, 137);
            this.label_GoogleSection.Name = "label_GoogleSection";
            this.label_GoogleSection.Size = new System.Drawing.Size(53, 15);
            this.label_GoogleSection.TabIndex = 1;
            this.label_GoogleSection.Text = "Google";
            // 
            // label_NotesSection
            // 
            this.label_NotesSection.AutoSize = true;
            this.label_NotesSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_NotesSection.Location = new System.Drawing.Point(15, 15);
            this.label_NotesSection.Name = "label_NotesSection";
            this.label_NotesSection.Size = new System.Drawing.Size(83, 15);
            this.label_NotesSection.TabIndex = 0;
            this.label_NotesSection.Text = "Lotus.Notes";
            // 
            // tabPage_Preferences
            // 
            this.tabPage_Preferences.AutoScroll = true;
            this.tabPage_Preferences.Controls.Add(this.checkBox_CreateNotification);
            this.tabPage_Preferences.Controls.Add(this.button1);
            this.tabPage_Preferences.Controls.Add(this.textBox_CustomDaysAhead);
            this.tabPage_Preferences.Controls.Add(this.checkBox_CustomDaysAhead);
            this.tabPage_Preferences.Controls.Add(this.checkBox_MinimizeToTray);
            this.tabPage_Preferences.Controls.Add(this.label_WindowSection);
            this.tabPage_Preferences.Controls.Add(this.checkBox_NotesServerAuth);
            this.tabPage_Preferences.Controls.Add(this.label_NotesConnectionSection);
            this.tabPage_Preferences.Controls.Add(this.checkBox_ConnectUsingSSL);
            this.tabPage_Preferences.Controls.Add(this.textBox_OtherCalName);
            this.tabPage_Preferences.Controls.Add(this.radioButton_OtherCalChoice);
            this.tabPage_Preferences.Controls.Add(this.radioButton_MainCalChoice);
            this.tabPage_Preferences.Controls.Add(this.label_CalendarSection);
            this.tabPage_Preferences.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage_Preferences.Location = new System.Drawing.Point(4, 22);
            this.tabPage_Preferences.Name = "tabPage_Preferences";
            this.tabPage_Preferences.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Preferences.Size = new System.Drawing.Size(572, 236);
            this.tabPage_Preferences.TabIndex = 1;
            this.tabPage_Preferences.Text = "Preferences";
            this.tabPage_Preferences.UseVisualStyleBackColor = true;
            // 
            // checkBox_CreateNotification
            // 
            this.checkBox_CreateNotification.AutoSize = true;
            this.checkBox_CreateNotification.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_CreateNotification.Location = new System.Drawing.Point(19, 99);
            this.checkBox_CreateNotification.Name = "checkBox_CreateNotification";
            this.checkBox_CreateNotification.Size = new System.Drawing.Size(215, 19);
            this.checkBox_CreateNotification.TabIndex = 12;
            this.checkBox_CreateNotification.Text = "Create Notifications (30 mins prior)";
            this.checkBox_CreateNotification.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 266);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 23);
            this.button1.TabIndex = 11;
            this.button1.Text = "Create Service";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // textBox_CustomDaysAhead
            // 
            this.textBox_CustomDaysAhead.Enabled = false;
            this.textBox_CustomDaysAhead.Location = new System.Drawing.Point(274, 177);
            this.textBox_CustomDaysAhead.Name = "textBox_CustomDaysAhead";
            this.textBox_CustomDaysAhead.Size = new System.Drawing.Size(60, 21);
            this.textBox_CustomDaysAhead.TabIndex = 10;
            this.textBox_CustomDaysAhead.Text = "14";
            this.textBox_CustomDaysAhead.LostFocus += new System.EventHandler(this.checkBox_CustomDaysAhead_LostFocus);
            // 
            // checkBox_CustomDaysAhead
            // 
            this.checkBox_CustomDaysAhead.AutoSize = true;
            this.checkBox_CustomDaysAhead.Location = new System.Drawing.Point(19, 179);
            this.checkBox_CustomDaysAhead.Name = "checkBox_CustomDaysAhead";
            this.checkBox_CustomDaysAhead.Size = new System.Drawing.Size(212, 19);
            this.checkBox_CustomDaysAhead.TabIndex = 9;
            this.checkBox_CustomDaysAhead.Text = "Configure Days Ahead (default 14)";
            this.checkBox_CustomDaysAhead.UseVisualStyleBackColor = true;
            this.checkBox_CustomDaysAhead.CheckedChanged += new System.EventHandler(this.checkBox_CustomDaysAhead_CheckedChanged);
            // 
            // checkBox_MinimizeToTray
            // 
            this.checkBox_MinimizeToTray.AutoSize = true;
            this.checkBox_MinimizeToTray.Location = new System.Drawing.Point(19, 235);
            this.checkBox_MinimizeToTray.Name = "checkBox_MinimizeToTray";
            this.checkBox_MinimizeToTray.Size = new System.Drawing.Size(120, 19);
            this.checkBox_MinimizeToTray.TabIndex = 8;
            this.checkBox_MinimizeToTray.Text = "Minimize To Tray";
            this.checkBox_MinimizeToTray.UseVisualStyleBackColor = true;
            // 
            // label_WindowSection
            // 
            this.label_WindowSection.AutoSize = true;
            this.label_WindowSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_WindowSection.Location = new System.Drawing.Point(16, 211);
            this.label_WindowSection.Name = "label_WindowSection";
            this.label_WindowSection.Size = new System.Drawing.Size(57, 15);
            this.label_WindowSection.TabIndex = 7;
            this.label_WindowSection.Text = "Window";
            // 
            // checkBox_NotesServerAuth
            // 
            this.checkBox_NotesServerAuth.AutoSize = true;
            this.checkBox_NotesServerAuth.Checked = true;
            this.checkBox_NotesServerAuth.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_NotesServerAuth.Location = new System.Drawing.Point(19, 153);
            this.checkBox_NotesServerAuth.Name = "checkBox_NotesServerAuth";
            this.checkBox_NotesServerAuth.Size = new System.Drawing.Size(157, 19);
            this.checkBox_NotesServerAuth.TabIndex = 6;
            this.checkBox_NotesServerAuth.Text = "Use server Athentication";
            this.checkBox_NotesServerAuth.UseVisualStyleBackColor = true;
            // 
            // label_NotesConnectionSection
            // 
            this.label_NotesConnectionSection.AutoSize = true;
            this.label_NotesConnectionSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_NotesConnectionSection.Location = new System.Drawing.Point(16, 128);
            this.label_NotesConnectionSection.Name = "label_NotesConnectionSection";
            this.label_NotesConnectionSection.Size = new System.Drawing.Size(83, 15);
            this.label_NotesConnectionSection.TabIndex = 5;
            this.label_NotesConnectionSection.Text = "Lotus Notes";
            // 
            // checkBox_ConnectUsingSSL
            // 
            this.checkBox_ConnectUsingSSL.AllowDrop = true;
            this.checkBox_ConnectUsingSSL.AutoSize = true;
            this.checkBox_ConnectUsingSSL.Checked = true;
            this.checkBox_ConnectUsingSSL.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_ConnectUsingSSL.Location = new System.Drawing.Point(19, 71);
            this.checkBox_ConnectUsingSSL.Name = "checkBox_ConnectUsingSSL";
            this.checkBox_ConnectUsingSSL.Size = new System.Drawing.Size(240, 19);
            this.checkBox_ConnectUsingSSL.TabIndex = 4;
            this.checkBox_ConnectUsingSSL.Text = "Use SSL to connect to Google calendar";
            this.checkBox_ConnectUsingSSL.UseVisualStyleBackColor = true;
            // 
            // textBox_OtherCalName
            // 
            this.textBox_OtherCalName.Enabled = false;
            this.textBox_OtherCalName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_OtherCalName.Location = new System.Drawing.Point(274, 44);
            this.textBox_OtherCalName.Name = "textBox_OtherCalName";
            this.textBox_OtherCalName.Size = new System.Drawing.Size(143, 21);
            this.textBox_OtherCalName.TabIndex = 3;
            this.textBox_OtherCalName.Text = "Lotus.Notes";
            this.textBox_OtherCalName.LostFocus += new System.EventHandler(this.textBox_OtherCalName_LostFocus);
            // 
            // radioButton_OtherCalChoice
            // 
            this.radioButton_OtherCalChoice.AutoSize = true;
            this.radioButton_OtherCalChoice.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_OtherCalChoice.Location = new System.Drawing.Point(160, 45);
            this.radioButton_OtherCalChoice.Name = "radioButton_OtherCalChoice";
            this.radioButton_OtherCalChoice.Size = new System.Drawing.Size(108, 19);
            this.radioButton_OtherCalChoice.TabIndex = 2;
            this.radioButton_OtherCalChoice.Text = "Other Calendar";
            this.radioButton_OtherCalChoice.UseVisualStyleBackColor = true;
            this.radioButton_OtherCalChoice.CheckedChanged += new System.EventHandler(this.radioButton_OtherCalChoice_CheckedChanged);
            // 
            // radioButton_MainCalChoice
            // 
            this.radioButton_MainCalChoice.AutoSize = true;
            this.radioButton_MainCalChoice.Checked = true;
            this.radioButton_MainCalChoice.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioButton_MainCalChoice.Location = new System.Drawing.Point(19, 43);
            this.radioButton_MainCalChoice.Name = "radioButton_MainCalChoice";
            this.radioButton_MainCalChoice.Size = new System.Drawing.Size(106, 19);
            this.radioButton_MainCalChoice.TabIndex = 1;
            this.radioButton_MainCalChoice.TabStop = true;
            this.radioButton_MainCalChoice.Text = "Main Calendar";
            this.radioButton_MainCalChoice.UseVisualStyleBackColor = true;
            this.radioButton_MainCalChoice.CheckedChanged += new System.EventHandler(this.radioButton_MainCalChoice_CheckedChanged);
            // 
            // label_CalendarSection
            // 
            this.label_CalendarSection.AutoSize = true;
            this.label_CalendarSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_CalendarSection.Location = new System.Drawing.Point(16, 18);
            this.label_CalendarSection.Name = "label_CalendarSection";
            this.label_CalendarSection.Size = new System.Drawing.Size(65, 15);
            this.label_CalendarSection.TabIndex = 0;
            this.label_CalendarSection.Text = "Calendar";
            // 
            // tabPage_SchedulingSync
            // 
            this.tabPage_SchedulingSync.Controls.Add(this.label_Minutes);
            this.tabPage_SchedulingSync.Controls.Add(this.textBox_ScheduleSync);
            this.tabPage_SchedulingSync.Controls.Add(this.checkBox_ScheduleSync);
            this.tabPage_SchedulingSync.Controls.Add(this.checkBox_SyncOnStartup);
            this.tabPage_SchedulingSync.Controls.Add(this.label_SchedulingSection);
            this.tabPage_SchedulingSync.Location = new System.Drawing.Point(4, 22);
            this.tabPage_SchedulingSync.Name = "tabPage_SchedulingSync";
            this.tabPage_SchedulingSync.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_SchedulingSync.Size = new System.Drawing.Size(572, 236);
            this.tabPage_SchedulingSync.TabIndex = 2;
            this.tabPage_SchedulingSync.Text = "Scheduling";
            this.tabPage_SchedulingSync.UseVisualStyleBackColor = true;
            // 
            // label_Minutes
            // 
            this.label_Minutes.AutoSize = true;
            this.label_Minutes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Minutes.Location = new System.Drawing.Point(172, 72);
            this.label_Minutes.Name = "label_Minutes";
            this.label_Minutes.Size = new System.Drawing.Size(51, 15);
            this.label_Minutes.TabIndex = 4;
            this.label_Minutes.Text = "minutes";
            // 
            // textBox_ScheduleSync
            // 
            this.textBox_ScheduleSync.Enabled = false;
            this.textBox_ScheduleSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_ScheduleSync.Location = new System.Drawing.Point(135, 69);
            this.textBox_ScheduleSync.Name = "textBox_ScheduleSync";
            this.textBox_ScheduleSync.Size = new System.Drawing.Size(31, 21);
            this.textBox_ScheduleSync.TabIndex = 3;
            this.textBox_ScheduleSync.Text = "30";
            // 
            // checkBox_ScheduleSync
            // 
            this.checkBox_ScheduleSync.AutoSize = true;
            this.checkBox_ScheduleSync.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_ScheduleSync.Location = new System.Drawing.Point(22, 71);
            this.checkBox_ScheduleSync.Name = "checkBox_ScheduleSync";
            this.checkBox_ScheduleSync.Size = new System.Drawing.Size(107, 19);
            this.checkBox_ScheduleSync.TabIndex = 2;
            this.checkBox_ScheduleSync.Text = "Schedule Sync";
            this.checkBox_ScheduleSync.UseVisualStyleBackColor = true;
            this.checkBox_ScheduleSync.CheckedChanged += new System.EventHandler(this.checkBox_ScheduleSync_CheckedChanged);
            // 
            // checkBox_SyncOnStartup
            // 
            this.checkBox_SyncOnStartup.AutoSize = true;
            this.checkBox_SyncOnStartup.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_SyncOnStartup.Location = new System.Drawing.Point(22, 44);
            this.checkBox_SyncOnStartup.Name = "checkBox_SyncOnStartup";
            this.checkBox_SyncOnStartup.Size = new System.Drawing.Size(111, 19);
            this.checkBox_SyncOnStartup.TabIndex = 1;
            this.checkBox_SyncOnStartup.Text = "Sync on Startup";
            this.checkBox_SyncOnStartup.UseVisualStyleBackColor = true;
            // 
            // label_SchedulingSection
            // 
            this.label_SchedulingSection.AutoSize = true;
            this.label_SchedulingSection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_SchedulingSection.Location = new System.Drawing.Point(19, 18);
            this.label_SchedulingSection.Name = "label_SchedulingSection";
            this.label_SchedulingSection.Size = new System.Drawing.Size(101, 15);
            this.label_SchedulingSection.TabIndex = 0;
            this.label_SchedulingSection.Text = "Schedule Sync";
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
            // 
            // label_Author
            // 
            this.label_Author.AutoSize = true;
            this.label_Author.Location = new System.Drawing.Point(333, 9);
            this.label_Author.Name = "label_Author";
            this.label_Author.Size = new System.Drawing.Size(41, 13);
            this.label_Author.TabIndex = 1;
            this.label_Author.Text = "Author:";
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
            // label_AuthorValue
            // 
            this.label_AuthorValue.AutoSize = true;
            this.label_AuthorValue.Location = new System.Drawing.Point(403, 9);
            this.label_AuthorValue.Name = "label_AuthorValue";
            this.label_AuthorValue.Size = new System.Drawing.Size(95, 13);
            this.label_AuthorValue.TabIndex = 4;
            this.label_AuthorValue.Text = "Kenneth Davidson";
            // 
            // label_VersionValue
            // 
            this.label_VersionValue.AutoSize = true;
            this.label_VersionValue.Location = new System.Drawing.Point(403, 31);
            this.label_VersionValue.Name = "label_VersionValue";
            this.label_VersionValue.Size = new System.Drawing.Size(22, 13);
            this.label_VersionValue.TabIndex = 5;
            this.label_VersionValue.Text = "0.1";
            // 
            // label_DateValue
            // 
            this.label_DateValue.AutoSize = true;
            this.label_DateValue.Location = new System.Drawing.Point(403, 53);
            this.label_DateValue.Name = "label_DateValue";
            this.label_DateValue.Size = new System.Drawing.Size(52, 13);
            this.label_DateValue.TabIndex = 6;
            this.label_DateValue.Text = "Mar 2010";
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
            this.textBox_AboutValue.Location = new System.Drawing.Point(403, 77);
            this.textBox_AboutValue.Multiline = true;
            this.textBox_AboutValue.Name = "textBox_AboutValue";
            this.textBox_AboutValue.Size = new System.Drawing.Size(180, 56);
            this.textBox_AboutValue.TabIndex = 9;
            this.textBox_AboutValue.Text = "Application based on iCalBridge (lngooglecalsync) ported to C#.";
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
            this.notifyIcon_Tray.BalloonTipText = "Notes to Google Sync";
            this.notifyIcon_Tray.BalloonTipTitle = "Sync";
            this.notifyIcon_Tray.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon_Tray.Icon")));
            this.notifyIcon_Tray.Text = "Notes To Google Sync";
            this.notifyIcon_Tray.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_Tray_MouseDoubleClick);
            // 
            // NotesToGoogleForm
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(602, 446);
            this.Controls.Add(this.panel_StatusPanel);
            this.Controls.Add(this.button_SaveSettings);
            this.Controls.Add(this.button_ManualSync);
            this.Controls.Add(this.textBox_AboutValue);
            this.Controls.Add(this.label_About);
            this.Controls.Add(this.pictureBox_Logo);
            this.Controls.Add(this.label_DateValue);
            this.Controls.Add(this.label_VersionValue);
            this.Controls.Add(this.label_AuthorValue);
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
            this.tabControl_MainControl.ResumeLayout(false);
            this.tabPage_Connect.ResumeLayout(false);
            this.tabPage_Connect.PerformLayout();
            this.tabPage_Preferences.ResumeLayout(false);
            this.tabPage_Preferences.PerformLayout();
            this.tabPage_SchedulingSync.ResumeLayout(false);
            this.tabPage_SchedulingSync.PerformLayout();
            this.tabPage_Debug.ResumeLayout(false);
            this.tabPage_Debug.PerformLayout();
            this.panel_StatusPanel.ResumeLayout(false);
            this.panel_StatusPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_Logo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private TabControl tabControl_MainControl;
        private TabPage tabPage_Connect;
        private PictureBox pictureBox_Logo;
        private TextBox textBox_GooglePassword;
        private TextBox textBox_GoogleLogin;
        private TabPage tabPage_Preferences;
        private Label label_Author;
        private Label label_Version;
        private Label label_Date;
        private Label label_AuthorValue;
        private Label label_VersionValue;
        private Label label_DateValue;
        private Label label_About;
        private TextBox textBox_AboutValue;
        private Button button_ManualSync;
        private Button button_SaveSettings;
        private Panel panel_StatusPanel;
        private ProgressBar progressBar_SyncProgress;
        private Label label_Status;
        private Label label_GoogleSection;
        private Label label_NotesSection;
        private Label label_NotesWebmailURL;
        private Label label_NotesPassword;
        private Label label_NotesLogin;
        private Label label_GooglePassword;
        private Label label_GoogleLogin;
        private TextBox textBox_NotesLogin;
        private TextBox textBox_WebmailURL;
        private Label label_CalendarSection;
        private RadioButton radioButton_MainCalChoice;
        private TextBox textBox_OtherCalName;
        private RadioButton radioButton_OtherCalChoice;
        private CheckBox checkBox_ConnectUsingSSL;
        private CheckBox checkBox_NotesServerAuth;
        private Label label_NotesConnectionSection;
        private TabPage tabPage_SchedulingSync;
        private CheckBox checkBox_SyncOnStartup;
        private Label label_SchedulingSection;
        private Label label_Minutes;
        private TextBox textBox_ScheduleSync;
        private CheckBox checkBox_ScheduleSync;
        private TextBox textBox_NotesPassword;
        private CheckBox checkBox_MinimizeToTray;
        private Label label_WindowSection;
        private TabPage tabPage_Debug;
        private TextBox textBox_Debug;
        private Button button_ClearDebug;
        private TextBox textBox_CustomDaysAhead;
        private CheckBox checkBox_CustomDaysAhead;
        private Button button1;
        private SyncPreferences config;
        private CheckBox checkBox_MessageBoxDebug;
        private CheckBox checkBox_CreateNotification;
        private BackgroundWorker backgroundWorker_SingleSync;
        private NotifyIcon notifyIcon_Tray;
        private IContainer components;
        private FormWindowState m_previousWindowState;

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
            notifyIcon_Tray.Visible = false;
            this.WindowState = m_previousWindowState;
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
        private void checkBox_CustomDaysAhead_LostFocus(object sender, EventArgs e)
        {
            try
            {
                int newDaysAhead = int.Parse(textBox_CustomDaysAhead.Text);
            }
            catch (FormatException ex)
            {
                PrintStringToDebug("Days Ahead: Value is not an appropriate number. \r\n Reverting to old value.");
                textBox_CustomDaysAhead.Text = config.GetPreference("DaysAhead");                
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

        private void textBox_OtherCalName_LostFocus(object sender, EventArgs e)
        {
            if (textBox_OtherCalName.Text == "")
            {
                textBox_OtherCalName.Text = "Lotus.Notes";
            }
        }

        private void checkBox_ScheduleSync_CheckedChanged(object sender, EventArgs e)
        {
            textBox_ScheduleSync.Enabled = checkBox_ScheduleSync.Checked;
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

            // Save Schedule Variables
            config.SetPreference("SyncOnStartup", checkBox_SyncOnStartup.Checked.ToString());
            config.SetPreference("ScheduleSync", checkBox_ScheduleSync.Checked.ToString());

            // Save Config File
            config.SavePreferences();

            // Confirm that file was saved
            PrintStringToDebug("Configuration was saved.");
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
        /// Method called when Manual Sync buttong has been pushed
        /// </summary>
        /// <param name="sender">Object creating the event</param>
        /// <param name="e">Event arguments</param>
        private void button_ManualSync_Click(object sender, EventArgs e)
        {
            // Disable the sync button, so we don't do this twice.
            button_ManualSync.Enabled = false;

            // Start the background worker.
            backgroundWorker_SingleSync.RunWorkerAsync();
        }

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
            }
        }

        // End of the NotesToGoogleAppClass
    }
}
