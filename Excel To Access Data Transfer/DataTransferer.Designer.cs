namespace Excel_To_Access_Data_Transfer
{
    partial class DataTransferer
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
            this.accessLocationTextbox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.resultsOutputTextbox = new System.Windows.Forms.TextBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnConnect = new System.Windows.Forms.Button();
            this.OpenDBDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnLoadDB = new System.Windows.Forms.Button();
            this.totaltextBox = new System.Windows.Forms.TextBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.movedtextBox = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnRevert = new System.Windows.Forms.Button();
            this.backgroundWorker2 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // accessLocationTextbox
            // 
            this.accessLocationTextbox.Location = new System.Drawing.Point(12, 25);
            this.accessLocationTextbox.Name = "accessLocationTextbox";
            this.accessLocationTextbox.Size = new System.Drawing.Size(417, 20);
            this.accessLocationTextbox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Access Database Location";
            // 
            // resultsOutputTextbox
            // 
            this.resultsOutputTextbox.Location = new System.Drawing.Point(12, 51);
            this.resultsOutputTextbox.Multiline = true;
            this.resultsOutputTextbox.Name = "resultsOutputTextbox";
            this.resultsOutputTextbox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.resultsOutputTextbox.Size = new System.Drawing.Size(528, 674);
            this.resultsOutputTextbox.TabIndex = 4;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(552, 106);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(105, 78);
            this.btnUpdate.TabIndex = 5;
            this.btnUpdate.Text = "Initiate Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(552, 50);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(105, 48);
            this.btnConnect.TabIndex = 6;
            this.btnConnect.Text = "Connect to Database";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // OpenDBDialog
            // 
            this.OpenDBDialog.FileName = "OpenDBDialog";
            // 
            // btnLoadDB
            // 
            this.btnLoadDB.Location = new System.Drawing.Point(435, 23);
            this.btnLoadDB.Name = "btnLoadDB";
            this.btnLoadDB.Size = new System.Drawing.Size(105, 23);
            this.btnLoadDB.TabIndex = 7;
            this.btnLoadDB.Text = "Browse Database";
            this.btnLoadDB.UseVisualStyleBackColor = true;
            this.btnLoadDB.Click += new System.EventHandler(this.btnLoadDB_Click);
            // 
            // totaltextBox
            // 
            this.totaltextBox.Location = new System.Drawing.Point(546, 323);
            this.totaltextBox.Name = "totaltextBox";
            this.totaltextBox.Size = new System.Drawing.Size(118, 20);
            this.totaltextBox.TabIndex = 9;
            this.totaltextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            // 
            // movedtextBox
            // 
            this.movedtextBox.Location = new System.Drawing.Point(546, 349);
            this.movedtextBox.Name = "movedtextBox";
            this.movedtextBox.Size = new System.Drawing.Size(118, 20);
            this.movedtextBox.TabIndex = 10;
            this.movedtextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(570, 492);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 54);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Export Update Report";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnRevert
            // 
            this.btnRevert.BackColor = System.Drawing.Color.Red;
            this.btnRevert.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnRevert.ForeColor = System.Drawing.SystemColors.InfoText;
            this.btnRevert.Location = new System.Drawing.Point(570, 652);
            this.btnRevert.Name = "btnRevert";
            this.btnRevert.Size = new System.Drawing.Size(75, 73);
            this.btnRevert.TabIndex = 13;
            this.btnRevert.Text = "REVERT \r\nFOLDER \r\nCLEANUP!";
            this.btnRevert.UseVisualStyleBackColor = false;
            this.btnRevert.Click += new System.EventHandler(this.btnRevert_Click);
            // 
            // backgroundWorker2
            // 
            this.backgroundWorker2.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker2_DoWork);
            // 
            // DataTransferer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(676, 737);
            this.Controls.Add(this.btnRevert);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.movedtextBox);
            this.Controls.Add(this.totaltextBox);
            this.Controls.Add(this.btnLoadDB);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.resultsOutputTextbox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.accessLocationTextbox);
            this.Name = "DataTransferer";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox accessLocationTextbox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox resultsOutputTextbox;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.OpenFileDialog OpenDBDialog;
        private System.Windows.Forms.Button btnLoadDB;
        private System.Windows.Forms.TextBox totaltextBox;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TextBox movedtextBox;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnRevert;
        private System.ComponentModel.BackgroundWorker backgroundWorker2;
    }
}

