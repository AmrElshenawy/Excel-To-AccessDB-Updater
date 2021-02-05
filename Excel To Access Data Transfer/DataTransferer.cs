using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
//using ExcelHST;
using XLHSTCommon;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Excel_To_Access_Data_Transfer
{
    public partial class DataTransferer : Form
    {
        string newLine = Environment.NewLine;   //Used to enter a new line
        public string accessLocation;           //Store connection string for the database      
        public string folderLocation;           //Store folder address containing files
        public string fileversion;              //Store read file version
        public string currentPart;              //Part that current index is referring to in for-loop
        public int rows;                        //Number of rows in LocalDataTable
        public int m;                           //Counter for number of files moved

        OleDbConnection dbConnection = new OleDbConnection();                   //Database connection object
        OleDbDataAdapter dataAdapter;                                           //Data adapter
        System.Data.DataTable LocalDataTable = new System.Data.DataTable();     //Initialize new data table

        public DataTransferer()
        {
            InitializeComponent();
            this.Text = "Excel to Access Database Updater";     //Window title
            btnConnect.Enabled = false;                         //------------------
            btnUpdate.Enabled = false;                          //  Disable buttons
            btnExport.Enabled = false;                          //------------------
        }

        private void btnLoadDB_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenDBDialog = new OpenFileDialog();                             //Open window to browse database location
            OpenDBDialog.Title = "Select a GlobaList.mdb Database";                         //Window title
            OpenDBDialog.Filter = "Access DB Files (*.mdb)|*.mdb|All Files (*.*)|*.*";      //Filter files ending in .mdb or All Files
            /* Once the user selects a file in the browse window 
             * the file name will be added to the Textbox for clarity.
             * Connected button will be enabled after */
            if (OpenDBDialog.ShowDialog() == DialogResult.OK)
            {
                string dbFileLocation = OpenDBDialog.FileName;
                accessLocationTextbox.AppendText(dbFileLocation);
            }
            btnConnect.Enabled = true;
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            //OleDB communication strings for connection
            string initialConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            accessLocation = initialConnection + accessLocationTextbox.Text;
            folderLocation = @"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\CURRENT\";
            dbConnection.ConnectionString = accessLocation;
            dbConnection.Open();
            dataAdapter = new OleDbDataAdapter("Select * From tblParts", dbConnection);     //Select the table "Parts" from Access DB
            dataAdapter.Fill(LocalDataTable);                                               //Fill the DataTable with data imported from Access DB
            rows = LocalDataTable.Rows.Count;

            //Confirm connection status to user
            if (dbConnection.State.Equals(ConnectionState.Open))
            {
                resultsOutputTextbox.AppendText("Connected to Database successfully!" + newLine);
                resultsOutputTextbox.AppendText("Total number of records to update: " + rows*2 + newLine + newLine);
                btnUpdate.Enabled = true;
            }
            else if(dbConnection.State.Equals(ConnectionState.Closed))
            {
                resultsOutputTextbox.AppendText("Failed to connect to Database." + newLine);
            }  
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            btnRevert.Enabled = false;
            backgroundWorker1.RunWorkerAsync();                         //Initiate a separate thread for perfoming the update and transfer
            backgroundWorker1.WorkerSupportsCancellation = true;
        }

        private void RefreshDBConnection()
        {
            if (dbConnection.State.Equals(ConnectionState.Open))
            {
                dbConnection.Close();
                LocalDataTable.Clear();
                btnConnect.PerformClick();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            /* 
             * Select PartFirstTestFileVersion from Parts table based on the PartFile column
             * Select PartfinalTestFileVersion from Parts table based on the PartFile column
             */ 
            OleDbCommand updateCommandTTA = new OleDbCommand("UPDATE tblParts SET [PartFirstTestFileVersion] = ? WHERE [PartFile] = ?", dbConnection);
            OleDbCommand updateCommandTTZ = new OleDbCommand("UPDATE tblParts SET [PartFinalTestFileVersion] = ? WHERE [PartFile] = ?", dbConnection);
            DirectoryInfo dir = new DirectoryInfo(folderLocation);
            m = 0;

            /*
             * Inititate a folder clean up before data transfer
             */ 
            this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
            {
                resultsOutputTextbox.AppendText("Running folder cleanup... " + newLine + newLine);
            }));
            FileInfo[] testfiles = dir.GetFiles("*", SearchOption.TopDirectoryOnly);                //Grab all files in directory and dump it in array
            for (int x = 0; x < testfiles.Length; x++)
            {
                bool found = false;

                /*
                 * For each file in testfiles[], scan through the database and find a matching file.
                 * If a matching file is found in the Database, turn flag on indicating a match.
                 */
                for (int y = 0; y < rows; y++)
                {
                    if (testfiles[x].ToString().ToUpper() == LocalDataTable.Rows[y].Field<string>(2) + ".TTA" || testfiles[x].ToString().ToUpper() == LocalDataTable.Rows[y].Field<string>(2) + ".TTZ")
                    {
                        found = true;
                        break;
                    }
                }
                /*
                 * If no match was found, move the file to a pre-set folder containing all unused files.
                 */
                if (found == false)
                {
                    string fileToMove = /*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\CURRENT\"*/ folderLocation + testfiles[x];
                    string moveTo = @"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\UNUSED TESTFILES\" /*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\"*/ + testfiles[x];
                    if (File.Exists(/*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\CURRENT\"*/ folderLocation + testfiles[x]))
                    {
                        File.Move(fileToMove, moveTo);
                        this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                        {
                            resultsOutputTextbox.AppendText("Found in folder: " + testfiles[x] + " --- NOT USED for a part in the Database. FILE MOVED!" + newLine);
                            movedtextBox.Clear();
                            movedtextBox.AppendText("Files Moved: " + ++m);
                        }));
                    }
                }
                /*
                 * Indicate to user that folder clean-up has been completed.
                 */
                if (testfiles.Length - x == 1)
                {
                    this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                    {
                        resultsOutputTextbox.AppendText(newLine + "Folder cleanup completed! " + newLine + newLine);
                        movedtextBox.Clear();
                        movedtextBox.AppendText("Files Moved: " + m);
                    }));
                }
            }
            for (int i = 0; i < rows; i++)
            {
                currentPart = LocalDataTable.Rows[i].Field<string>(2);
                this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                {
                    resultsOutputTextbox.AppendText("Current Part: " + currentPart + newLine);
                }));
                #region Commented out
                /*FileInfo[] testfiles = dir.GetFiles(/*currentPart.Substring(0, currentPart.Length-5) + "*"folderLocation, SearchOption.TopDirectoryOnly);
                for (int x = 0; x < testfiles.Length; x++)
                {
                    bool found = false;
                    for (int y = 0; y < rows; y++)
                    {
                        if (testfiles[x].ToString().ToUpper() == LocalDataTable.Rows[y].Field<string>(2) + ".TTA" || testfiles[x].ToString().ToUpper() == LocalDataTable.Rows[y].Field<string>(2) + ".TTZ")
                        {
                            found = true;
                            break;
                        }
                    }
                    if (found == false)
                    {
                        string fileToMove = @"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\CURRENT\" + testfiles[x];
                        string moveTo = @"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\" + testfiles[x];
                        if(File.Exists(@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\CURRENT\" + testfiles[x]))
                        {
                            File.Move(fileToMove, moveTo);
                            this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                            {
                                resultsOutputTextbox.AppendText("Found in folder: " + testfiles[x] + " --- NOT USED for a part in the Database. FILE MOVED!" + newLine);
                                movedtextBox.Clear();
                                movedtextBox.AppendText("Files Moved: " + ++m);
                            }));
                        }     
                    }
                }*/
                #endregion

                /*
                 * Scan through the Database, if no matching file found in the folder,
                 * Indicate that the file version in N/A for both TTA and TTZ.
                 */
                if (!File.Exists(folderLocation + currentPart + ".TTA"))       
                {                                                                                                   
                    updateCommandTTA.Parameters.Clear();
                    updateCommandTTA.Parameters.AddWithValue("@PartFirstTestFileVersion", "N/A");
                    updateCommandTTA.Parameters.AddWithValue("@PartFile", currentPart);
                    updateCommandTTA.ExecuteNonQuery();
                    updateCommandTTA.Parameters.Clear();
                }
                if (!File.Exists(folderLocation + currentPart + ".TTZ"))
                {
                    updateCommandTTZ.Parameters.Clear();
                    updateCommandTTZ.Parameters.AddWithValue("@PartFinalTestFileVersion", "N/A");
                    updateCommandTTZ.Parameters.AddWithValue("@PartFile", currentPart);
                    updateCommandTTZ.ExecuteNonQuery();
                    updateCommandTTZ.Parameters.Clear();
                }

                /*
                 * Otherwise, if a file is found in the folder for a matching part in the Database,
                 * Depending on the file type TTA or TTZ, open the file and read the file version.
                 * Insert the file version in the apporpiate column for that PartFile.
                 */
                for (int j = 0; j < testfiles.Length; j++)
                {
                    if (testfiles[j].ToString().ToUpper() == (currentPart + ".TTA"))
                    {
                        WinHSTTestFile HSTFile = new WinHSTTestFile(folderLocation + (currentPart + ".TTA"), true);
                        fileversion = HSTFile.TestParameters.FileVersion;
                        updateCommandTTA.Parameters.Clear();
                        updateCommandTTA.Parameters.AddWithValue("@PartFirstTestFileVersion", fileversion);
                        updateCommandTTA.Parameters.AddWithValue("@PartFile", currentPart);
                        updateCommandTTA.ExecuteNonQuery();
                        this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                        {
                            resultsOutputTextbox.AppendText("Updated: " + currentPart + " FIRST using " + (currentPart + ".TTA") + newLine);
                        }));
                        updateCommandTTA.Parameters.Clear();
                    }
                    if (testfiles[j].ToString().ToUpper() == (currentPart + ".TTZ"))
                    {
                        WinHSTTestFile HSTFile = new WinHSTTestFile(folderLocation + (currentPart + ".TTZ"), true);
                        fileversion = HSTFile.TestParameters.FileVersion;
                        updateCommandTTZ.Parameters.Clear();
                        updateCommandTTZ.Parameters.AddWithValue("@PartFinalTestFileVersion", fileversion);
                        updateCommandTTZ.Parameters.AddWithValue("@PartFile", currentPart);
                        updateCommandTTZ.ExecuteNonQuery();
                        this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                        {
                            resultsOutputTextbox.AppendText("Updated: " + currentPart + " FINAL using " + (currentPart + ".TTZ") + newLine + newLine);
                        }));
                        updateCommandTTZ.Parameters.Clear();
                    }
                }

                //Progress Counter
                this.totaltextBox.Invoke(new MethodInvoker(delegate ()
                {
                    totaltextBox.Clear();
                    totaltextBox.AppendText(((i * 2) + 2) + "/" + rows * 2 + " Completed!");
                }));

                //Indicate that the update has been completed
                if (rows - i == 1)
                {
                    this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                    {
                        resultsOutputTextbox.AppendText("Update completed!");
                        btnExport.Enabled = true;
                    }));
                    dbConnection.Close();
                    backgroundWorker1.CancelAsync();
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            DirectoryInfo reportDir = new DirectoryInfo(@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\EADU Exported Database Update Reports");
            FileInfo[] Reports = reportDir.GetFiles("*.txt", SearchOption.TopDirectoryOnly).OrderBy(f => f.CreationTime).ToArray();         //Sort exported reports by date of creation
            string latestReport = string.Empty;
            string subreportName = string.Empty;

            /*
             * Extract the relevant piece of the report name string to find the ID# of the report
             * Once the latest ID# has been recognized, increment the report number for the current report instance to be exported.
             */
            if (Reports.Length != 0)
            {
                latestReport = Reports[Reports.Length - 1].ToString();
                subreportName = latestReport.Substring(1, 5);
            }
            string reportNumberString = string.Empty;
            int reportNumber = 0;
            for (int i = 0; i < subreportName.Length; i++)
            {
                if (Char.IsDigit(subreportName[i]))
                {
                    reportNumberString += subreportName[i];
                }
            }
            if (reportNumberString.Length > 0)
            {
                reportNumber = Int32.Parse(reportNumberString);
            }
            TextWriter txt = new StreamWriter(@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\EADU Exported Database Update Reports\#" + ++reportNumber + " Database Update Completed - " + m + " Files Removed" + ".txt");
            txt.Write(resultsOutputTextbox.Text);
            txt.Close();
            MessageBox.Show("Report Exported to:" + newLine + newLine + @"...\HST\TESTFILE\MASTERS\EADU Exported Database Update Reports");
        }

        private void btnRevert_Click(object sender, EventArgs e)
        {
            backgroundWorker2.RunWorkerAsync();
            backgroundWorker2.WorkerSupportsCancellation = true;
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            DirectoryInfo revertDir = new DirectoryInfo(/*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\"*/@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\UNUSED TESTFILES\");
            FileInfo[] revertFiles = revertDir.GetFiles("*", SearchOption.TopDirectoryOnly);

            /*
             *  In the event that an unused file needs to be moved back to main folder, initiate a folder clean-up revert
             */
            bool isEmpty = !Directory.EnumerateFiles(/*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\"*/@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\UNUSED TESTFILES\").Any();

            /*
             * Check if there are any folders to revert
             */
            if (isEmpty)
            {
                this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                {
                    resultsOutputTextbox.AppendText("There is nothing to revert." + newLine);
                    backgroundWorker2.CancelAsync();
                }));
            }
            else
            {
                this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                {
                    resultsOutputTextbox.AppendText("Reverting folder cleanup..." + newLine);
                }));
            }

            /*
             * Grab every single folder in the unused files folder and move it back to the main folder to complete the revert.
             */
            for (int i = 0; i < revertFiles.Length; i++)
            {
                string filetoMove =/*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\"*/@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\UNUSED TESTFILES\" + revertFiles[i];
                string moveTo = /*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\CURRENT\"*/@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\CURRENT\" + revertFiles[i];
                if (File.Exists(/*@"G:\Test Eng\Amr Access\TestFiles Cleanup Experiment\Unused TestFiles\"*/@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\UNUSED TESTFILES\" + revertFiles[i]))
                {
                    File.Move(filetoMove, moveTo);
                    this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                    {
                        resultsOutputTextbox.Clear();
                        resultsOutputTextbox.AppendText("Reverting folder cleanup..." + newLine);
                        resultsOutputTextbox.AppendText((i + 1) + "/" + revertFiles.Length + newLine);
                    }));
                }

                //Indicate to user that the folder revert has been completed.
                if (revertFiles.Length - i == 1)
                {
                    this.resultsOutputTextbox.Invoke(new MethodInvoker(delegate ()
                    {
                        resultsOutputTextbox.AppendText("Folder Cleanup Revert Completed." + newLine + newLine);
                    }));
                    
                    backgroundWorker2.CancelAsync();
                }
            }
        }
    }
}