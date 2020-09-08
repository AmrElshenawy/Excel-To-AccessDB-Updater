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
            this.label2 = new System.Windows.Forms.Label();
            this.excelLocation = new System.Windows.Forms.TextBox();
            this.resultsOutputTextbox = new System.Windows.Forms.TextBox();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnConnect = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // accessLocationTextbox
            // 
            this.accessLocationTextbox.Location = new System.Drawing.Point(153, 12);
            this.accessLocationTextbox.Name = "accessLocationTextbox";
            this.accessLocationTextbox.Size = new System.Drawing.Size(511, 20);
            this.accessLocationTextbox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Access Database Location";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Excel File Location";
            // 
            // excelLocation
            // 
            this.excelLocation.Location = new System.Drawing.Point(153, 41);
            this.excelLocation.Name = "excelLocation";
            this.excelLocation.Size = new System.Drawing.Size(511, 20);
            this.excelLocation.TabIndex = 3;
            // 
            // resultsOutputTextbox
            // 
            this.resultsOutputTextbox.Location = new System.Drawing.Point(12, 83);
            this.resultsOutputTextbox.Multiline = true;
            this.resultsOutputTextbox.Name = "resultsOutputTextbox";
            this.resultsOutputTextbox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.resultsOutputTextbox.Size = new System.Drawing.Size(516, 221);
            this.resultsOutputTextbox.TabIndex = 4;
            // 
            // btnUpdate
            // 
            this.btnUpdate.Location = new System.Drawing.Point(589, 135);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(149, 87);
            this.btnUpdate.TabIndex = 5;
            this.btnUpdate.Text = "Start Update";
            this.btnUpdate.UseVisualStyleBackColor = true;
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(670, 10);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(118, 23);
            this.btnConnect.TabIndex = 6;
            this.btnConnect.Text = "Connect to Database";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // DataTransferer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 316);
            this.Controls.Add(this.btnConnect);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.resultsOutputTextbox);
            this.Controls.Add(this.excelLocation);
            this.Controls.Add(this.label2);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox excelLocation;
        private System.Windows.Forms.TextBox resultsOutputTextbox;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnConnect;
    }
}

