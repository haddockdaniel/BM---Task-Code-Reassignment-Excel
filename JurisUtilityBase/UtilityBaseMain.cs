using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public List<Row> badRows = new List<Row>();

        string pathToExcelFile = "";

        public Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

        Workbook xlWorkbook;

        _Worksheet xlWorksheet;

        bool codeIsNumeric = false;

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
            DataSet ds = displayErrors();
            ReportDisplay rpds = new ReportDisplay(ds);
            rpds.Show();

        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            if (!string.IsNullOrEmpty(pathToExcelFile))
            {
                toolStripStatusLabel.Text = "Running. Please Wait...";
                xlApp.Visible = false;
                xlWorkbook = xlApp.Workbooks.Open(pathToExcelFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                Row currentRow = null;

                for (int a = 2; a <= rowCount; a++)
                {
                    Microsoft.Office.Interop.Excel.Range range1 = xlWorksheet.Rows[a]; //For all columns in rows
                    //Range range1 = worksheet.Columns[1]; //for all rows in column 1

                    int col = 1;
                    currentRow = new Row(); //custom row class. I only use a few of the attributes but they are all programmed in the class if needed
                    foreach (Range r in range1.Cells) //range1.Cells represents all the columns/rows
                    {
                        
                        if (col == 1)
                            currentRow.client = Convert.ToString(r.Value).Trim();
                        else if (col == 2)
                            currentRow.oldTask = Convert.ToString(r.Value).Trim();
                        else if (col == 3)
                            currentRow.newTask = Convert.ToString(r.Value).Trim();
                        else if (col > 3)
                            break;
                        col++;
                    }

                    //if there IS an error (returns true)
                    if (checkClientAndTaskCodes(currentRow))
                    {
                        badRows.Add(currentRow);
                    }
                    else // no error so continue
                    {
                            //do sql stuff here
                            string SQL = "Update timeentry Set taskcode='" + currentRow.newTask.Trim() + "' where ClientSysNbr = " + currentRow.clisys.ToString() + " and taskcode='" + currentRow.oldTask.Trim() + "' and EntryStatus < 8";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            SQL = "Update T Set tbdtaskcd='" + currentRow.newTask.Trim() + "' from Timebatchdetail as T inner join matter as M on M.matsysnbr = T.TBDMatter " +
                            "  inner join client as C on C.clisysnbr = M.matclinbr inner join unbilledtime as U on T.TBDBATCH = U.uTBATCH and T.TBDRECNBR = U.uTRECNBR where tbdtaskcd='" + currentRow.oldTask.Trim() + "' and C.clisysnbr = " + currentRow.clisys.ToString();
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                            SQL = "Update U Set uttaskcd='" + currentRow.newTask.Trim() + "' from UnbilledTime U inner join matter as M on M.matsysnbr = U.utmatter " +
                            " inner join client as C on C.clisysnbr = M.matclinbr where C.clisysnbr =" + currentRow.clisys.ToString() + " and uttaskcd= '" + currentRow.oldTask.Trim() + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    }
                }

                UpdateStatus("All Task codes updated.", 1, 1);
                toolStripStatusLabel.Text = "Status: Ready to Execute";

                if (badRows.Count == 0)
                    MessageBox.Show("The process is complete without error", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                else
                {
                    DialogResult rr = MessageBox.Show("The process is complete but there were" + "\r\n" + "errors. Would you like to see them?", "Errors", MessageBoxButtons.YesNo, MessageBoxIcon.None);
                    if (rr == DialogResult.Yes)
                    {
                        DataSet ds = displayErrors();
                        ReportDisplay rpds = new ReportDisplay(ds);
                        rpds.Show();
                    }
                }
                xlWorkbook.Close(0);
                xlApp.Quit();

                while (Marshal.ReleaseComObject(xlApp) != 0) { }
                while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
                while (Marshal.ReleaseComObject(xlWorksheet) != 0) { }
                xlApp = null;
                xlWorkbook = null;
                xlWorksheet = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Environment.Exit(0);
            }
            else
                MessageBox.Show("Please browse to your Excel file first", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            
        }

        private void getNumberSettings()
        {
            string sql = "  select SpTxtValue from sysparam where SpName = 'FldClient'";
            DataSet dds = _jurisUtility.RecordsetFromSQL(sql);
            string cell = "";
            if (dds != null && dds.Tables.Count > 0)
            {
                foreach (DataRow dr in dds.Tables[0].Rows)
                    cell = dr[0].ToString();
            }

            string[] test = cell.Split(',');


            if (test[1].Equals("C"))
                codeIsNumeric = false;
            else
                codeIsNumeric = true;




        }


        private string formatClientCode(string code)
        {
            getNumberSettings();
            string formattedCode = "";
            if (codeIsNumeric)
            {
                formattedCode = "000000000000" + code;
                formattedCode = formattedCode.Substring(formattedCode.Length - 12, 12);
            }
            else
                formattedCode = code;
            return formattedCode;

        }

        //returns false if EID exists in timeentry, timebatchdetail and unbilledtime table as well as the taskcode existing in taskcode, otherwise returns true
        //which means at least one of these tests were failed and they need to be fixed
        private bool checkClientAndTaskCodes(Row currRow)
        {
            DataSet ds1;
            //old taskcode
            string SQL = "Select * from taskcode where TaskCdCode = '" + currRow.oldTask.Trim() + "'";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The old task code is not valid";
                return true;
            }
            ds1.Clear();

            //new taskcode
            SQL = "Select * from taskcode where TaskCdCode = '" + currRow.newTask.Trim() + "'";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The new task code is not valid";
                return true;
            }
            ds1.Clear();

            //client
            currRow.client = formatClientCode(currRow.client.Trim());
            SQL = "select clisysnbr from client where clicode = '" + currRow.client.Trim() + "'";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1 != null && ds1.Tables.Count > 0)
            {
                foreach (DataRow dr in ds1.Tables[0].Rows)
                    currRow.clisys = Convert.ToInt32(dr[0].ToString());
            }
            else
            {
                currRow.error = "The client code is not valid";
                return true;
            }
            ds1.Clear();


            //only reachable if all of these sql queries return at least one row (they should all only return 1 row btw) :)
            return false;
        }



        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            if (xlWorkbook != null)
                xlWorkbook.Close(0);
            if (xlApp != null)
                xlApp.Quit();
            if (xlApp != null)
                while (Marshal.ReleaseComObject(xlApp) != 0) { }
            if (xlWorkbook != null)
                while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
            if (xlWorksheet != null)
                while (Marshal.ReleaseComObject(xlWorksheet) != 0) { }
            xlApp = null;
            xlWorkbook = null;
            xlWorksheet = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            System.Environment.Exit(0);
          
        }

        private DataSet displayErrors()
        {
            DataSet ds = new DataSet();
            System.Data.DataTable errorTable = ds.Tables.Add("Errors");
            errorTable.Columns.Add("ClientCode");
            errorTable.Columns.Add("OldTaskCode");
            errorTable.Columns.Add("NewTaskCode");
            errorTable.Columns.Add("Error");

            string err = "";
            for (int a = 1; a < 5; a++)
            {
                
                DataRow errorRow = ds.Tables["Errors"].NewRow();
                errorRow["ClientCode"] = "000" + a.ToString();
                errorRow["OldTaskCode"] = "L10" + a.ToString();
                errorRow["NewTaskCode"] = "L11" + a.ToString();
                errorRow["Error"] = err;
                ds.Tables["Errors"].Rows.Add(errorRow);
            }


            //  foreach (Row r in badRows)
            //  {
            //      DataRow errorRow = ds.Tables["Errors"].NewRow();
            //      errorRow["ClientCode"] = r.client;
            //      errorRow["OldTaskCode"] = r.oldTask;
            //     errorRow["NewTaskCode"] = r.newTask;
            //     errorRow["Error"] = r.error;
            //      ds.Tables["Errors"].Rows.Add(errorRow);
            //  }

            return ds;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".xls" || Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".xlsx" || Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".xlsm")
                {
                    pathToExcelFile = openFileDialog1.FileName;
                    label2.Text = "File Chosen: " + Path.GetFileName(pathToExcelFile);

                }
                else
                    MessageBox.Show("Only valid Excel files can be seleced (.xls, .xlsx, .xlsm)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            } 
        }




    }
}
