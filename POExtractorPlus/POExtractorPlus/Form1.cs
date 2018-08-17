using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace POExtractorPlus
{
    public partial class Form1 : Form
    {
        public Controller Controller { get; set; }
        public string[] POFiles { get; set; }

        private bool IsLicenced { get; set; }

        public Form1()
        {
            InitializeComponent();
            IsLicenced = false;
            backgroundWorker1.DoWork += BackgroundWorker1_DoWork;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;

            Init();
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Visible = false;

            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Invoke(new MethodInvoker(delegate ()
            {
                if (ValidateForm())
                {
                    List<string> rejectedFiles = this.Controller.Extract(
                     accountTypeComboBox.Text,
                     this.POFiles,
                     destinationTextBox.Text);

                    if (rejectedFiles.Count > 0)
                    {
                        if (this.Height != 405)
                        {
                            this.Height = 405;
                            this.panel1.Visible = true;
                        }

                        this.rejectedPOFilesListBox.DataSource = rejectedFiles;
                    }
                }
            }));
        }

        private void Init()
        {
            string defaultDestination = string.Format(@"{0}\{1}", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "PO Plus Docs");
            this.destinationTextBox.Text = defaultDestination;
            this.Controller = new Controller();

            CheckTrialPeriodAtFirstTime();
            SetTrialPeriodTimer();
        }

        private void SetTrialPeriodTimer()
        {
            Thread timerThread = new Thread(CheckTrailPeriod);
            timerThread.Start();
        }

        private void CheckTrailPeriod()
        {
            //CheckTrialPeriod();
            System.Timers.Timer timer = new System.Timers.Timer();
            // timer.Interval = 43200000;
            timer.Interval = 10000;
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
        }

        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            CheckTrialPeriod();
        }

        private void CheckTrialPeriod()
        {
            this.Invoke(new MethodInvoker(delegate ()
            {
                //DateTime dt = new DateTime(2018, 7, 25);
                //if (DateTime.Now > dt)
                //{
                //    IsLicenced = false;
                //    trailVersionTextBox.Text = "Expired. Please contact the administrator.";
                //    button3.Enabled = false;
                //}
                //else
                //{
                //    IsLicenced = true;
                //    var remainingDays = (dt - DateTime.Now).Days;
                //    string text = string.Format("Trail Version. Expire in {0} days.", remainingDays);
                //    if (remainingDays == 0)
                //    {
                //        text = "Trail Version. Expire today.";
                //    }

                //    trailVersionTextBox.Text = text;
                //}

                if (CheckForInternetConnection())
                {
                    string jsonString = new WebClient().DownloadString("https://raw.githubusercontent.com/hbtmrt/POExtractorPlus/master/POExtractorPlus/onlinedata.json");
                    bool isForceExpire = jsonString.Contains("true");

                    if (isForceExpire)
                    {
                        IsLicenced = false;
                        trailVersionTextBox.Text = "Expired. Please contact the administrator.";
                        button3.Enabled = false;
                    }
                    else
                    {
                        DateTime dt = new DateTime(2018, 8, 18);
                        if (DateTime.Now > dt)
                        {
                            IsLicenced = false;
                            trailVersionTextBox.Text = "Expired. Please contact the administrator.";
                            button3.Enabled = false;
                        }
                        else
                        {
                            IsLicenced = true;
                            var remainingDays = (dt - DateTime.Now).Days;
                            string text = string.Format("Trail Version. Expire in {0} days.", remainingDays);
                            if (remainingDays == 0)
                            {
                                text = "Trail Version. Expire today.";
                            }
                            button3.Enabled = true;
                            trailVersionTextBox.Text = text;
                        }
                    }
                }
                else
                {
                    IsLicenced = false;
                    trailVersionTextBox.Text = "Please connect to the internet";
                    button3.Enabled = false;
                }
            }));
        }

        private static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                using (client.OpenRead("http://clients3.google.com/generate_204"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private void CheckTrialPeriodAtFirstTime()
        {
            //DateTime dt = new DateTime(2018, 7, 25);
            //if (DateTime.Now > dt)
            //{
            //    IsLicenced = false;
            //    trailVersionTextBox.Text = "Expired. Please contact the administrator.";
            //    button3.Enabled = false;
            //}
            //else
            //{
            //    IsLicenced = true;
            //    var remainingDays = (dt - DateTime.Now).Days;
            //    string text = string.Format("Trail Version. Expire in {0} days.", remainingDays);
            //    if (remainingDays == 0)
            //    {
            //        text = "Trail Version. Expire today.";
            //    }

            //    trailVersionTextBox.Text = text;
            //}

            if (CheckForInternetConnection())
            {
                string jsonString = new WebClient().DownloadString("https://raw.githubusercontent.com/hbtmrt/POExtractorPlus/master/POExtractorPlus/onlinedata.json");
                bool isForceExpire = jsonString.Contains("true");

                if (isForceExpire)
                {
                    IsLicenced = false;
                    trailVersionTextBox.Text = "Expired. Please contact the administrator.";
                    button3.Enabled = false;
                }
                else
                {
                    DateTime dt = new DateTime(2018, 8, 18);
                    if (DateTime.Now > dt)
                    {
                        IsLicenced = false;
                        trailVersionTextBox.Text = "Expired. Please contact the administrator.";
                        button3.Enabled = false;
                    }
                    else
                    {
                        IsLicenced = true;
                        var remainingDays = (dt - DateTime.Now).Days;
                        string text = string.Format("Trail Version. Expire in {0} days.", remainingDays);
                        if (remainingDays == 0)
                        {
                            text = "Trail Version. Expire today.";
                        }
                        button3.Enabled = true;
                        trailVersionTextBox.Text = text;
                    }
                }
            }
            else
            {
                IsLicenced = false;
                trailVersionTextBox.Text = "Please connect to the internet";
                button3.Enabled = false;
            }
        }

        private void OnBrowsingDestination(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    this.destinationTextBox.Text = fbd.SelectedPath;
                    Properties.Settings.Default.Destination = fbd.SelectedPath;
                }
            }
        }

        private void OnBrowingPOFiles(object sender, EventArgs e)
        {
            using (var pfd = new OpenFileDialog())
            {
                pfd.Multiselect = true;
                pfd.Filter = "Pdf Files|*.pdf";

                DialogResult result = pfd.ShowDialog();

                if (result == DialogResult.OK)
                {
                    if (pfd.FileNames.Length > 0)
                    {
                        if (!IsLicenced && pfd.FileNames.Length > 3)
                        {
                            MessageBox.Show("You can select at most 3 files in the trial version.");
                            this.poFilesTextBox.Text = "";
                            SetSelectedFilesToListBox(null);
                            return;
                        }

                        this.POFiles = pfd.FileNames;
                        this.poFilesTextBox.Text = string.Join(",", pfd.FileNames);
                        SetSelectedFilesToListBox(pfd.FileNames);
                    }
                }
            }
        }

        private void SetSelectedFilesToListBox(string[] files)
        {
            //string[] fileNamesOnly = new string[files.Length];
            List<string> fileNamesOnly = new List<string>();

            if (files != null)
            {
                for (int i = 0; i < files.Length; i++)
                {
                    fileNamesOnly.Add(Path.GetFileName(files[0]));
                    //fileNamesOnly[i] = Path.GetFileName(files[0]);
                }
            }

            this.allPOFilesListBox.DataSource = fileNamesOnly;
        }

        private void OnExtracting(object sender, EventArgs e)
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
        }

        private bool ValidateForm()
        {
            bool validated = true;
            string errorMessage = string.Empty;

            if (accountTypeComboBox.SelectedItem == null)
            {
                validated = false;
                errorMessage = string.Format("{0} \n{1}", errorMessage, "The account type should be selected.");
            }

            if (this.POFiles == null || (this.POFiles != null && this.POFiles.Length == 0))
            {
                validated = false;
                errorMessage = string.Format("{0} \n{1}", errorMessage, "At least one PO should be selected.");
            }

            if (!validated)
            {
                MessageBox.Show(errorMessage);
            }

            return validated;
        }
    }
}
