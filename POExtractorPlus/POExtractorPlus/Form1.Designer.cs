namespace POExtractorPlus
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.accountTypeComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.poFilesTextBox = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.destinationTextBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.allPOFilesListBox = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.rejectedPOFilesListBox = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.windowTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Account Type";
            // 
            // accountTypeComboBox
            // 
            this.accountTypeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.accountTypeComboBox.FormattingEnabled = true;
            this.accountTypeComboBox.Items.AddRange(new object[] {
            "BRAZIL",
            "LACL-Mexico",
            "EU-CANADA-USA-APD"});
            this.accountTypeComboBox.Location = new System.Drawing.Point(109, 11);
            this.accountTypeComboBox.Name = "accountTypeComboBox";
            this.accountTypeComboBox.Size = new System.Drawing.Size(331, 21);
            this.accountTypeComboBox.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "PO Files";
            // 
            // poFilesTextBox
            // 
            this.poFilesTextBox.Enabled = false;
            this.poFilesTextBox.Location = new System.Drawing.Point(109, 42);
            this.poFilesTextBox.Name = "poFilesTextBox";
            this.poFilesTextBox.Size = new System.Drawing.Size(331, 20);
            this.poFilesTextBox.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(458, 40);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.OnBrowingPOFiles);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(60, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Destination";
            // 
            // destinationTextBox
            // 
            this.destinationTextBox.Enabled = false;
            this.destinationTextBox.Location = new System.Drawing.Point(109, 70);
            this.destinationTextBox.Name = "destinationTextBox";
            this.destinationTextBox.Size = new System.Drawing.Size(331, 20);
            this.destinationTextBox.TabIndex = 6;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(458, 67);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "Browse";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.OnBrowsingDestination);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(109, 96);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "Extract";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.OnExtracting);
            // 
            // allPOFilesListBox
            // 
            this.allPOFilesListBox.FormattingEnabled = true;
            this.allPOFilesListBox.Location = new System.Drawing.Point(109, 146);
            this.allPOFilesListBox.Name = "allPOFilesListBox";
            this.allPOFilesListBox.ScrollAlwaysVisible = true;
            this.allPOFilesListBox.Size = new System.Drawing.Size(424, 95);
            this.allPOFilesListBox.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 146);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(91, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Selected PO Files";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 13);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Rejected PO Files";
            // 
            // rejectedPOFilesListBox
            // 
            this.rejectedPOFilesListBox.FormattingEnabled = true;
            this.rejectedPOFilesListBox.Location = new System.Drawing.Point(97, 13);
            this.rejectedPOFilesListBox.Name = "rejectedPOFilesListBox";
            this.rejectedPOFilesListBox.ScrollAlwaysVisible = true;
            this.rejectedPOFilesListBox.Size = new System.Drawing.Size(424, 95);
            this.rejectedPOFilesListBox.TabIndex = 12;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.rejectedPOFilesListBox);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(12, 247);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(532, 117);
            this.panel1.TabIndex = 13;
            this.panel1.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(547, 251);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.allPOFilesListBox);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.destinationTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.poFilesTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.accountTypeComboBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PO Extractor Plus";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox accountTypeComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox poFilesTextBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox destinationTextBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ListBox allPOFilesListBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ListBox rejectedPOFilesListBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ToolTip windowTooltip;
    }
}

