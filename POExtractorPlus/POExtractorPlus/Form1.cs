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
        public Form1()
        {
            InitializeComponent();
            Init();
        }

        private void Init()
        {
            string defaultDestination = Properties.Settings.Default.Destination;
            this.destinationTextBox.Text = defaultDestination;
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
                    if (pfd.FileNames.Length > 0) {
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
    }
}
