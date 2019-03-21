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

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
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

            string sql = "select vencode as VendorCode, venname as VendorName, vengets1099 as VenGets1099 from Vendor where venactive = 'Y' and vensysnbr not in (1,2)";
            DataSet vendors = _jurisUtility.RecordsetFromSQL(sql);
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = vendors.Tables[0]; // dataset
            dataGridView1.Columns[0].Width = 90;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 90;

            sql = "select case when venpaymentgroup<='' or venpaymentgroup is null then 'NO GROUP' else VENPAYMENTGROUP end as VendorGroups from vendor group by case when venpaymentgroup<='' or venpaymentgroup is null then 'NO GROUP' else VENPAYMENTGROUP end";
            DataSet vendors1 = _jurisUtility.RecordsetFromSQL(sql);
            dataGridView2.AutoGenerateColumns = true;
            dataGridView2.DataSource = vendors1.Tables[0]; // dataset
            dataGridView2.Columns[0].Width = 250;


        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);
            string SQL = "";
            if (rbInactivate.Checked) //they want to inactivate
            {
                if (rbByDate.Checked) //by date
                {
                    DateTime dt = new DateTime();
                    dt = dateTimePicker1.Value;
                    SQL = "Update vendor set VenActive = 'N' where vensysnbr not in (1,2) and vensysnbr in (Select vensysnbr from vendor where venactive = 'Y') and  vensysnbr in (Select vchvendor from voucher group by vchvendor having Max(vchinvoicedate) <= convert(datetime,'" + dt.ToString("yyyyMMdd") + "', 101)) and vensysnbr not in  ( Select vchvendor from voucher where vchstatus <>'P') ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    SQL = "Update vendor set VenActive = 'N' where vensysnbr in (" +
                    "select distinct vensysnbr from vendor " +
                    " left outer  join checkregister on CkRegVend = vensysnbr" +
                    " where vensysnbr in (Select vensysnbr from vendor where venactive = 'Y') " +
                    " and vensysnbr not in  ( Select vchvendor from voucher where vchstatus <>'P') " +
                    " and ckregdate is not null " +
                    " and CkRegVend is not null " +
                    " group by vensysnbr " +
                    " having max(convert(datetime,ckregdate, 101)) <= convert(datetime,'" + dt.ToString("yyyyMMdd") + "', 101))";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    UpdateStatus("Updating Vendors...", 1, 2);



                    List<int> convertedVendors = new List<int>();
                    SQL = "select distinct vensysnbr from vendor where  convert(datetime,VenLastPurchaseDate, 101) = convert(datetime, '19000101', 101)";
                    DataSet converted = _jurisUtility.RecordsetFromSQL(SQL);
                    foreach (DataRow row in converted.Tables[0].Rows)
                        convertedVendors.Add(Int32.Parse(row["vensysnbr"].ToString()));

                    SQL = "select vensysnbr from Vendor_Log where RecordType = 1 group by vensysnbr having Max([DateTimeStamp]) <= convert(datetime,'" + dt.ToString("yyyyMMdd") + "', 101)";
                    DataSet log = _jurisUtility.RecordsetFromSQL(SQL);

                    List<int> tempList = convertedVendors.ToList();
                    foreach (int sysnbr in convertedVendors)
                    {
                        foreach (DataRow row in log.Tables[0].Rows)
                        {
                            if (sysnbr == Int32.Parse(row["vensysnbr"].ToString()))
                            {
                                var itemToRemove = convertedVendors.Single(r => r == sysnbr);
                                tempList.Remove(itemToRemove);
                            }
                        }
                    }

                    string IDs = String.Join(",", tempList);

                    SQL = "Update vendor set VenActive = 'N' where vensysnbr not in (1,2) and vensysnbr in (" + IDs + ") and vensysnbr not in  ( Select vchvendor from voucher where vchstatus <>'P') ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    SQL = "delete from documenttree where  DTDocClass = 7000 and DTKeyL not in (1,2) and DTKeyL in (Select vensysnbr from vendor WHERE VenActive = 'N')";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                    UpdateStatus("Updating DocTree...", 2, 2);

                   
                    SQL = "Select [VenCode] As VendorCode,[VenName] as VendorName,[VenAddress] as Address,[VenCity] as City,[VenState] as State,[VenZip] ZipCode ,[VenCountry] as Country from vendor where vensysnbr in ( Select vchvendor from voucher where vchstatus <>'P') and  vensysnbr in (Select vchvendor from voucher   group by vchvendor having Max(vchinvoicedate) <= convert(datetime,'" + dt.ToString("yyyyMMdd") + "',101)) ";
                    DataSet ss = _jurisUtility.RecordsetFromSQL(SQL);
                    if (ss.Tables[0].Rows.Count != 0)
                    {
                        DialogResult d = MessageBox.Show("Would you like to see the vendors who were unable to be closed due to vouchers?", "Exception report", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (d == System.Windows.Forms.DialogResult.Yes)
                        {
                            ReportDisplay s1 = new ReportDisplay(ss);
                            s1.Show();
                        }
                    }
                    else
                        MessageBox.Show("The process finished with no exceptions", "Process complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else // by list selection
                {
                    string venList = "";
                    foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    {
                        venList = venList + "'" + r.Cells[0].Value.ToString() + "',";
                    }
                    venList = venList.TrimEnd(',');
                    SQL = "Update vendor set VenActive = 'N' where vensysnbr not in (1,2)and vensysnbr in (Select vensysnbr from vendor where vencode in ( " + venList + ")) and vensysnbr not in  ( Select vchvendor from voucher where vchstatus <>'P') ";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    UpdateStatus("Updating Vendors...", 1, 2);
                    SQL = "delete from documenttree where  DTDocClass = 7000 and DTKeyL not in (1,2) and DTKeyL in (Select vensysnbr from vendor WHERE VenActive = 'N')";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    UpdateStatus("Updating DocTree...", 2, 2);
                    SQL = "Select [VenCode] As VendorCode,[VenName] as VendorName,[VenAddress] as Address,[VenCity] as City,[VenState] as State,[VenZip] ZipCode ,[VenCountry] as Country from vendor where vencode in ( " + venList + ") and vensysnbr in  ( Select vchvendor from voucher where vchstatus <>'P')  ";
                    DataSet ss = _jurisUtility.RecordsetFromSQL(SQL);

                    if (ss.Tables[0].Rows.Count != 0)
                    {
                        DialogResult d = MessageBox.Show("Would you like to see the vendors who were unable to be closed due to vouchers?", "Exception report", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (d == System.Windows.Forms.DialogResult.Yes)
                        {
                            ReportDisplay s1 = new ReportDisplay(ss);
                            s1.Show();
                        }
                    }
                    else
                        MessageBox.Show("The process finished with no exceptions", "Process complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            else if (rb1099.Checked)
            {
                string type = "";
                if (comboBox1.SelectedIndex > -1)
                {
                    type = comboBox1.SelectedItem.ToString().Split(' ')[0];
                    SQL = "update vendor set vengets1099 = 'Y', ven1099box = " + type + "  where venpaymentgroup in (";
                    foreach (DataGridViewRow r in dataGridView2.SelectedRows)
                    {
                        SQL = SQL + "'" + r.Cells[0].Value.ToString() + "',";
                    }
                    SQL = SQL.TrimEnd(',');
                    SQL = SQL.Replace("NO GROUP", "");
                    SQL = SQL + ")";
                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    UpdateStatus("Updating Vendors...", 1, 1);
                    MessageBox.Show("The process finished with no exceptions", "Process complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                    MessageBox.Show("Please select a new 1099 Type from the drop down");

            }

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
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
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
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
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
                //generates output of the report for before and after the change will be made to client
                string SQLTkpr = getReportSQL();
                if (string.IsNullOrEmpty(SQLTkpr))
                    MessageBox.Show("Please select the proper options before running a report");
                else
                {
                    DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                    ReportDisplay rpds = new ReportDisplay(myRSTkpr);
                    rpds.Show();
                }

 
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            if (rbInactivate.Checked)
            {
                if (rbByName.Checked) //selection of vendor
                {
                    reportSQL = "select [VenCode] As VendorCode,[VenName] as VendorName,[VenAddress] as Address,[VenCity] as City,[VenState] as State,[VenZip] ZipCode ,[VenCountry] as Country from vendor where vencode in (";
                    foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                    {
                        reportSQL = reportSQL + "'" + r.Cells[0].Value.ToString() + "',";
                    }
                    reportSQL = reportSQL.TrimEnd(',');
                    reportSQL = reportSQL + ")";
                }
                else //by date
                {
                    DateTime dt = new DateTime();
                    dt = dateTimePicker1.Value;
                    reportSQL = "select [VenCode] As VendorCode,[VenName] as VendorName,[VenAddress] as Address,[VenCity] as City,[VenState] as State,[VenZip] ZipCode ,[VenCountry] as Country from vendor where vensysnbr in (Select vchvendor from voucher   group by vchvendor having Max(vchinvoicedate) <= '" + dt.ToString("yyyyMMdd") + "') and venactive = 'Y' and vensysnbr not in (1,2)";

                }
            }
            else if (rb1099.Checked)
            {
                string type = "";
                if (comboBox1.SelectedIndex > -1)
                {
                    type = comboBox1.SelectedItem.ToString();
                    reportSQL = "select [VenCode] As VendorCode,[VenName] as VendorName,[VenAddress] as Address,[VenCity] as City,[VenState] as State,[VenZip] ZipCode ,[VenCountry] as Country, '" + type + "' as New1099Type, venpaymentgroup as GroupName from vendor where vensysnbr not in (1,2) and venpaymentgroup in (";
                    foreach (DataGridViewRow r in dataGridView2.SelectedRows)
                    {
                        reportSQL = reportSQL + "'" + r.Cells[0].Value.ToString() + "',";
                    }
                    reportSQL = reportSQL.TrimEnd(',');
                    reportSQL = reportSQL.Replace("NO GROUP", "");
                    reportSQL = reportSQL + ")";
                }

            }
            return reportSQL;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbInactivate.Checked)
            { 
                groupBox2.Visible = false;
                groupBox1.Visible = true;
            }
            else if (rb1099.Checked)
            {
                groupBox1.Visible = false;
                groupBox2.Visible = true;
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (rb1099.Checked)
            { 
                groupBox1.Visible = false;
                groupBox2.Visible = true;
            }
            else if (rbInactivate.Checked)
            {
                groupBox2.Visible = false;
                groupBox1.Visible = true;
            }

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
    }
}
