using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace POExtractorPlus
{
    public partial class Form1 : Form
    {
        public Controller Controller { get; set; }
        public string[] POFiles { get; set; }

        public Form1()
        {
            InitializeComponent();
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
            string defaultDestination = Properties.Settings.Default.Destination;
            this.destinationTextBox.Text = defaultDestination;
            this.Controller = new Controller();
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
                        this.POFiles = pfd.FileNames;
                        this.poFilesTextBox.Text = string.Join(",", pfd.FileNames);
                        SetSelectedFilesToListBox(pfd.FileNames);
                    }
                }
            }
        }

        private void SetSelectedFilesToListBox(string[] files)
        {
            string[] fileNamesOnly = new string[files.Length];
            for (int i = 0; i < files.Length; i++)
            {
                fileNamesOnly[i] = Path.GetFileName(files[0]);
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
