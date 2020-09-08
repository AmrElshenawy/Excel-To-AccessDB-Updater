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
using ExcelHST;
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
        string newLine = Environment.NewLine;
        public string accessLocation;
        public string folderLocation;
        public string cell;
        public string fileversion;

        OleDbConnection dbConnection = new OleDbConnection();
        OleDbDataAdapter dataAdapter;
        System.Data.DataTable LocalDataTable = new System.Data.DataTable();

        int rowPosition = 0;
        int rowNumber = 0;

        public DataTransferer()
        {
            InitializeComponent();
            this.Text = "Excel to Access Data Transferer";
            //string initialConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            //string accessLocation = initialConnection + accessLocationTextbox.Text;
            accessLocation = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\aelshenawy\Desktop\GlobalList.mdb";
            //folderLocation = @"C:\Users\aelshenawy\Desktop\Testfiles Test";
            folderLocation = @"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\CURRENT";

        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            dbConnection.ConnectionString = accessLocation;
            dbConnection.Open();
            dataAdapter = new OleDbDataAdapter("Select * From tblParts", dbConnection);
            dataAdapter.Fill(LocalDataTable);

            int rows = LocalDataTable.Rows.Count;
            //int columns = LocalDataTable.Columns.Count;

            DirectoryInfo dir = new DirectoryInfo(folderLocation);
            FileInfo[] testfiles = dir.GetFiles("GJ215155*", SearchOption.TopDirectoryOnly);
            resultsOutputTextbox.AppendText("IN FOLDER: " + testfiles[0] + newLine);
            resultsOutputTextbox.AppendText("IN FOLDER: " + testfiles[1] + newLine);

            foreach (var item in testfiles)
            {
                if (item.ToString() == "GJ215155.tta" || item.ToString() == "GJ215155.ttz")
                {
                    resultsOutputTextbox.AppendText("Found: " + item + newLine);
                    WinHSTTestFile HSTFile = new WinHSTTestFile(@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\CURRENT\" + item.ToString(), true);
                    fileversion = HSTFile.TestParameters.FileVersion;
                    resultsOutputTextbox.AppendText("File Version: " + fileversion + newLine);

                    //ThisAddIn.LoadHSTfileToNewWorkbook(@"\\mk-testeng\TestEng\HSTDATA\HST\TESTFILE\MASTERS\CURRENT\GJ215155.tta");
                }
            }

            for (int i = 0; i < rows; i++)
            {
                if(LocalDataTable.Rows[i].Field<string>(2) == "GJ215155")
                {
                    LocalDataTable.Rows[i]["PartFirstTestFileVersion"] = fileversion;
                }
                //cell = LocalDataTable.Rows[i]["PartFile"].ToString();
                //resultsOutputTextbox.AppendText(cell + newLine);
            }

            OleDbCommand updateCommand = new OleDbCommand("UPDATE tblParts SET [PartFirstTestFileVersion] = ?, [PartFinalTestFileVersion] = ? WHERE [PartFile] = ?", dbConnection);
            updateCommand.Parameters.AddWithValue("@PartFirstTestFileVersion", fileversion);
            updateCommand.Parameters.AddWithValue("@PartFinalTestFileVersion", fileversion);
            updateCommand.Parameters.AddWithValue("@PartFile", "GJ215155");
            updateCommand.ExecuteNonQuery();
        }
    }
}
