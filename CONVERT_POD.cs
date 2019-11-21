using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using Syncfusion.XlsIO;

namespace DCH_CONVERT_POD
{
    public partial class CONVERT_POD : Form
    {
        public CONVERT_POD()
        {
            InitializeComponent();
        }


        public class EPOD
        {
            public string CO { get; set; }
            public string Doc_T { get; set; }
            public string Doc_N { get; set; }
            public string Inv_D { get; set; }
            public string Due_D { get; set; }
            public string Cust { get; set; }
            public string Ord_T { get; set; }
            public string Ord_N { get; set; }
            public string BP { get; set; }
            public string Ref { get; set; }
            public string Sls { get; set; }
            public string C08 { get; set; }
            public string Term { get; set; }
            public string Net { get; set; }
            public string Amt_E { get; set; }
            public string VAT { get; set; }
            public string Amt_I { get; set; }
            public string O_Amt { get; set; }
            public string Actl_D { get; set; }
            public string Rcpt_D { get; set; }
            public string Import_D { get; set; }
        }

        private DataTable AddColumnDatatable()
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Columns.Add("CO");
                dt.Columns.Add("Doc_T");
                dt.Columns.Add("Doc_N");
                dt.Columns.Add("Inv_D");
                dt.Columns.Add("Due_D");
                dt.Columns.Add("Cust");
                dt.Columns.Add("Ord_T");
                dt.Columns.Add("Ord_N");
                dt.Columns.Add("BP");
                dt.Columns.Add("Ref");
                dt.Columns.Add("Sls");
                dt.Columns.Add("C08");
                dt.Columns.Add("Term");
                dt.Columns.Add("Net");
                dt.Columns.Add("Amt_E");
                dt.Columns.Add("VAT");
                dt.Columns.Add("Amt_I");
                dt.Columns.Add("O_Amt");
                dt.Columns.Add("Actl_D");
                dt.Columns.Add("Rcpt_D");
                dt.Columns.Add("Import_D");
                return dt;
            }
            catch (Exception ex)
            {
                return dt;
            }
        }

        private DataTable AddColumnDatatable_TRAN()
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Columns.Add("CO");
                dt.Columns.Add("Doc_T");
                dt.Columns.Add("Doc_N");
                dt.Columns.Add("Trn_D");
                dt.Columns.Add("Act_C");
                dt.Columns.Add("Sys_D");
                dt.Columns.Add("F_Dep");
                dt.Columns.Add("T_Dep");
                dt.Columns.Add("Doc_C");
                dt.Columns.Add("Rea_C");
                dt.Columns.Add("Rem");
                dt.Columns.Add("Pay_T");
                dt.Columns.Add("Date1");
                dt.Columns.Add("Date2");
                dt.Columns.Add("Fdate");
                dt.Columns.Add("Tdate");
                dt.Columns.Add("Pay_Amt");
                dt.Columns.Add("Import_D");
 
                return dt;
            }
            catch (Exception ex)
            {
                return dt;
            }
        }

        private DataTable AddColumnDatatable_AMT()
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Columns.Add("Doc_T");
                dt.Columns.Add("Doc_N");
                dt.Columns.Add("O_Amt");
                return dt;
            }
            catch (Exception ex)
            {
                return dt;
            }
        }
        public DataTable ReadExcel(string fileName, string fileExt, string sPrefix, out bool bResult,out string sError)
        {
            DataTable dtexcel = new DataTable();
            DataTable dt = new DataTable();
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            //int i;
            bResult = false;
            sError = string.Empty;
            try
            {
                string conn = string.Empty;
                if (fileExt.CompareTo(".xls") == 0)
                    conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                else
                    conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                using (OleDbConnection con = new OleDbConnection(conn))
                {
                    try
                    {

                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                        dtexcel = new DataTable();
                        oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                        dt = new DataTable();
                        if (sPrefix.ToUpper() == "ITF-POD-HEAD")
                        {
                            dt = AddColumnDatatable();
                            for (int i = 1; i < dtexcel.Rows.Count; i++)
                            {
                                DataRow dr = dt.NewRow();
                                dr["CO"] = dtexcel.Rows[i]["F1"].ToString();
                                dr["Doc_T"] = dtexcel.Rows[i]["F2"].ToString();
                                dr["Doc_N"] = dtexcel.Rows[i]["F3"].ToString();
                                if (dtexcel.Rows[i]["F4"].ToString().Trim() != "" && dtexcel.Rows[i]["F4"].ToString().Trim() != "-")
                                {
                                    dr["Inv_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F4"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                }
                                else
                                {
                                    dr["Inv_D"] = "01/01/50 00:00:00";
                                }

                                if (dtexcel.Rows[i]["F5"].ToString().Trim() != "")
                                {
                                    dr["Due_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F5"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                }
                                else
                                {
                                    dr["Due_D"] = "01/01/50 00:00:00";
                                }

                                dr["Cust"] = dtexcel.Rows[i]["F6"].ToString();
                                dr["Ord_T"] = dtexcel.Rows[i]["F7"].ToString();
                                dr["Ord_N"] = dtexcel.Rows[i]["F8"].ToString();
                                dr["BP"] = dtexcel.Rows[i]["F9"].ToString();
                                dr["Ref"] = dtexcel.Rows[i]["F10"].ToString();
                                dr["Sls"] = dtexcel.Rows[i]["F11"].ToString();
                                dr["C08"] = dtexcel.Rows[i]["F12"].ToString();
                                dr["Term"] = dtexcel.Rows[i]["F13"].ToString();
                                dr["Net"] = dtexcel.Rows[i]["F14"].ToString();
                                dr["Amt_E"] = dtexcel.Rows[i]["F15"].ToString();
                                dr["VAT"] = dtexcel.Rows[i]["F16"].ToString();
                                dr["Amt_I"] = dtexcel.Rows[i]["F17"].ToString();
                                dr["O_Amt"] = dtexcel.Rows[i]["F18"].ToString();
                                if (dtexcel.Rows[i]["F19"].ToString().Trim() != "" && dtexcel.Rows[i]["F19"].ToString().Trim() != "-")
                                {
                                    dr["Actl_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F19"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                }
                                else
                                {
                                    dr["Actl_D"] = "01/01/50 00:00:00";
                                }

                                if (dtexcel.Rows[i]["F20"].ToString().Trim() != "" && dtexcel.Rows[i]["F20"].ToString().Trim() != "-")
                                {
                                    dr["Rcpt_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F20"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                }
                                else
                                {
                                    dr["Rcpt_D"] = "01/01/50 00:00:00";
                                }

                                if (dtexcel.Rows[i]["F21"].ToString().Trim() != "" && dtexcel.Rows[i]["F21"].ToString().Trim() != "-")
                                {
                                    dr["Import_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F21"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                }
                                else
                                {
                                    dr["Import_D"] = "01/01/50 00:00:00";
                                }


                                dt.Rows.Add(dr);
                            }
                            
                        }
                        else
                        {
                            if (sPrefix.ToUpper() == "ITF-POD-TRAN")
                            {
                                dt = AddColumnDatatable_TRAN();
                                for (int i = 1; i < dtexcel.Rows.Count; i++)
                                {

                                    DataRow dr = dt.NewRow();
                                    dr["CO"] = dtexcel.Rows[i]["F1"].ToString();
                                    dr["Doc_T"] = dtexcel.Rows[i]["F2"].ToString();
                                    dr["Doc_N"] = dtexcel.Rows[i]["F3"].ToString();
                                    if (dtexcel.Rows[i]["F4"].ToString().Trim() != "" && dtexcel.Rows[i]["F4"].ToString().Trim() != "-")
                                    {
                                        dr["Trn_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F4"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                    }
                                    else
                                    {
                                        dr["Trn_D"] = "01/01/50 00:00:00";
                                    }
                                    //dr["Trn_D"] = dtexcel.Rows[i]["F4"].ToString();
                                    dr["Act_C"] = dtexcel.Rows[i]["F5"].ToString();
                                    if (dtexcel.Rows[i]["F6"].ToString().Trim() != "" && dtexcel.Rows[i]["F6"].ToString().Trim() != "-")
                                    {
                                        dr["Sys_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F6"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                    }
                                    else
                                    {
                                        dr["Sys_D"] = "01/01/50 00:00:00";
                                    }
                                    //dr["Sys_D"] = dtexcel.Rows[i]["F6"].ToString();
                                    dr["F_Dep"] = dtexcel.Rows[i]["F7"].ToString();
                                    dr["T_Dep"] = dtexcel.Rows[i]["F8"].ToString();
                                    dr["Doc_C"] = dtexcel.Rows[i]["F9"].ToString();
                                    dr["Rea_C"] = dtexcel.Rows[i]["F10"].ToString();
                                    dr["Rem"] = dtexcel.Rows[i]["F11"].ToString();
                                    dr["Pay_T"] = dtexcel.Rows[i]["F12"].ToString();
                                    dr["Date1"] = dtexcel.Rows[i]["F13"].ToString();
                                    dr["Date2"] = dtexcel.Rows[i]["F14"].ToString();
                                    dr["Fdate"] = dtexcel.Rows[i]["F15"].ToString();
                                    dr["Tdate"] = dtexcel.Rows[i]["F16"].ToString();
                                    dr["Pay_Amt"] = dtexcel.Rows[i]["F17"].ToString();
                                    if (dtexcel.Rows[i]["F18"].ToString().Trim() != "" && dtexcel.Rows[i]["F18"].ToString().Trim() != "-")
                                    {
                                        dr["Import_D"] = Convert.ToDateTime(dtexcel.Rows[i]["F18"].ToString()).ToString("dd/MM/yy hh:mm:ss");
                                    }
                                    else
                                    {
                                        dr["Import_D"] = "01/01/50 00:00:00";
                                    }
                                    //dr["Import_D"] = dtexcel.Rows[i]["F18"].ToString();


                                    dt.Rows.Add(dr);
                                }
                            }
                            else
                            {
                                if (sPrefix.ToUpper() == "ITF-UPDO-AMT")
                                {
                                    dt = AddColumnDatatable_AMT();
                                    for (int i = 1; i < dtexcel.Rows.Count; i++)
                                    {
                                        DataRow dr = dt.NewRow();
                                        dr["Doc_T"] = dtexcel.Rows[i]["F1"].ToString();
                                        dr["Doc_N"] = dtexcel.Rows[i]["F2"].ToString();
                                        dr["O_Amt"] = dtexcel.Rows[i]["F3"].ToString();


                                        dt.Rows.Add(dr);
                                    }
                                }
                            }
                        }


                    }
                    catch (Exception ex)
                    {
                        sError = ex.Message.ToString();
                        //MessageBox.Show(sError);
                        bResult = false;
                        return dt;
                    }
                }
                bResult = true;
                return dt;
            }
            catch (Exception ex)
            {
                sError = ex.Message.ToString();
                bResult = false;
                return dt;
            }
            finally
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }

        }

        private void WriteDataToFile(DataTable submittedDataTable, string submittedFilePath)
        {
            try
            {
                int i = 0;
                StreamWriter sw = null;
                StringBuilder result = new StringBuilder();
                sw = new StreamWriter(submittedFilePath, false);
                for (i = 0; i < submittedDataTable.Columns.Count - 1; i++)
                {

                    //sw.Write(submittedDataTable.Columns[i].ColumnName + @"\t");
                    result.Append(submittedDataTable.Columns[i].ColumnName);
                    result.Append("\t"); // tab delimited
                }
                //sw.Write(submittedDataTable.Columns[i].ColumnName);
                result.Append(submittedDataTable.Columns[i].ColumnName);
                //sw.WriteLine();
                result.AppendLine();

                foreach (DataRow row in submittedDataTable.Rows)
                {
                    object[] array = row.ItemArray;

                    for (i = 0; i < array.Length - 1; i++)
                    {
                        //sw.Write(array[i].ToString() + @"\t");
                        result.Append(array[i].ToString());
                        result.Append("\t"); // tab delimited
                    }
                    //sw.Write(array[i].ToString());
                    result.Append(array[i].ToString());
                    //sw.WriteLine();
                    result.AppendLine();

                }
                sw.Write(result.ToString());
                sw.Close();
            }
            catch (Exception)
            {
                return;
            }
            
        }
        private void CONVERT_POD_Load(object sender, EventArgs e)
        {
            string PATH_SOURCE = System.Configuration.ConfigurationSettings.AppSettings["PATH_SOURCE"].ToString();
            string PATH_TARGET = System.Configuration.ConfigurationSettings.AppSettings["PATH_TARGET"].ToString();
            string PATH_BAK = System.Configuration.ConfigurationSettings.AppSettings["PATH_BAK"].ToString();
            DataTable dt_output = new DataTable();
            StringBuilder result = new StringBuilder();
            string HEAD = System.Configuration.ConfigurationSettings.AppSettings["HEAD"].ToString();
            string TRAN = System.Configuration.ConfigurationSettings.AppSettings["TRAN"].ToString();
            string sAMT = System.Configuration.ConfigurationSettings.AppSettings["AMT"].ToString();
            string PATH_Error = System.Configuration.ConfigurationSettings.AppSettings["PATH_Error"].ToString();

            string MailFrom = System.Configuration.ConfigurationSettings.AppSettings["MailFrom"].ToString();
            string MailTo = System.Configuration.ConfigurationSettings.AppSettings["MailTo"].ToString();
            string smtp = System.Configuration.ConfigurationSettings.AppSettings["SMTP"].ToString();
            string strSubject = System.Configuration.ConfigurationSettings.AppSettings["sSubjectmail"].ToString();
            ClassLibrarySendMail.ClassLibrarySendMail classmail = new ClassLibrarySendMail.ClassLibrarySendMail();
            try
            {
                // Header
                bool bResult;
                string sError = string.Empty;
                string result_Move;
                foreach (string filepath in Directory.GetFiles(PATH_SOURCE, HEAD + "*.xlsx"))
                {
                    bResult = false;
                    dt_output = new DataTable();
                    dt_output = ReadExcel(filepath, ".xlsx", HEAD, out bResult, out sError);
                    result_Move = Path.GetFileName(filepath);
                    if (bResult)
                    {
                        #region Create Text file
                        string FileName = Path.GetFileNameWithoutExtension(filepath);
                        WriteDataToFile(dt_output, PATH_TARGET + @"\\" + FileName + ".txt");
                        #endregion

                        #region Move File
                       
                      

                        string sPath_bak = PATH_BAK + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_bak);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_bak);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_bak + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);
                        classmail.Sendmail(MailTo, smtp, "Success Convert POD xlsx to Text Header ", MailFrom, "Success Convert POD");
                        #endregion
                    }
                    else // Send email error
                    {
                        string sPath_error = PATH_Error + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_error);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_error);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_error + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);

                        classmail.Sendmail(MailTo, smtp, "Error Convert POD xlsx to Text Header :" + sError.ToString(), MailFrom, strSubject);
                    }


                    System.Threading.Thread.Sleep(1000);
                }

                // Transaction
                foreach (string filepath in Directory.GetFiles(PATH_SOURCE, TRAN + "*.xlsx"))
                {
                    bResult = false;
                    dt_output = new DataTable();
                    dt_output = ReadExcel(filepath, ".xlsx", TRAN, out bResult, out sError);
                    result_Move = Path.GetFileName(filepath);
                    if (bResult)
                    {
                        #region Create Text file
                        string FileName = Path.GetFileNameWithoutExtension(filepath);
                        WriteDataToFile(dt_output, PATH_TARGET + @"\\" + FileName + ".txt");
                        #endregion

                        #region Move File
   
                       

                        string sPath_bak = PATH_BAK + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_bak);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_bak);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_bak + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);
                        classmail.Sendmail(MailTo, smtp, "Success Convert POD xlsx to Text Transaction ", MailFrom, "Success Convert POD");
                        #endregion
                    }
                    else // Send email error
                    {
                        string sPath_error = PATH_Error + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_error);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_error);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_error + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);

                        classmail.Sendmail(MailTo, smtp, "Error Convert POD xlsx to Text Transaction :" + sError.ToString(), MailFrom, strSubject);
                    }

                    System.Threading.Thread.Sleep(1000);
                }

                // AMT
                foreach (string filepath in Directory.GetFiles(PATH_SOURCE, sAMT + "*.xlsx"))
                {
                    bResult = false;
                    dt_output = new DataTable();
                    dt_output = ReadExcel(filepath, ".xlsx", sAMT, out bResult, out sError);
                    result_Move = Path.GetFileName(filepath);
                    if (bResult)
                    {
                        #region Create Text file
                        string FileName = Path.GetFileNameWithoutExtension(filepath);
                        WriteDataToFile(dt_output, PATH_TARGET + @"\\" + FileName + ".txt");
                        #endregion

                        #region Move File
                       
                       

                        string sPath_bak = PATH_BAK + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_bak);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_bak);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_bak + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);
                        classmail.Sendmail(MailTo, smtp, "Success Convert POD xlsx to Text AMT ", MailFrom, "Success Convert POD");
                        #endregion
                    }
                    else // Send email error
                    {
                        string sPath_error = PATH_Error + "/" + DateTime.Now.ToString("yyyyMMdd");
                        bool exists = System.IO.Directory.Exists(sPath_error);
                        if (!exists)
                        {
                            System.IO.Directory.CreateDirectory(sPath_error);
                        }

                        System.IO.File.Move(PATH_SOURCE + "/" + result_Move, sPath_error + "/" + DateTime.Now.ToString("yyyyMMddhhmmss", (new System.Globalization.CultureInfo("en-US"))) + "_" + result_Move);

                        classmail.Sendmail(MailTo, smtp, "Error Convert POD xlsx to Text AMT :" + sError.ToString(), MailFrom, strSubject);
                    }

                    System.Threading.Thread.Sleep(1000);
                }

            }
            catch (Exception ex)
            {
                return;
            }
            finally
            {
                this.Dispose();
            }
        }
    }
}
