using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Windows.Forms;
using System.Linq;
using System.Configuration;
using System.Collections;
using System.Data.SqlClient;
using System.IO;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows;
using CrystalDecisions.CrystalReports.Engine;


namespace CrystalReportsApplication1
{
    public partial class Form1 : Form
    {
        ReportDocument DeviationReportView = new ReportDocument();
        clsDBAccess objDbAccess = new clsDBAccess();
        Connection objCon = new Connection();
        Hashtable parameters = new Hashtable();
        String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
        String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();
        DataSet dsAgency = new DataSet();

        bool StopLoop = true;

        public Form1()
        {
            InitializeComponent();
        }



        protected void ApplyLoginInfo()
        {
            ConnectionInfo crConnectionInfo = new ConnectionInfo();

            crConnectionInfo.ServerName = ConfigurationManager.AppSettings["DBServerName"];
            crConnectionInfo.UserID = ConfigurationManager.AppSettings["DBLoginInfo"];
            crConnectionInfo.Password = ConfigurationManager.AppSettings["DBPwdInfo"];

            Database crDatabase = DeviationReportView.Database;
            Tables crTables = crDatabase.Tables;
            CrystalDecisions.CrystalReports.Engine.Table crTable;
            TableLogOnInfo crTableLogOnInfo;


            //Loop through all tables in the report and apply the
            //connection information for each table.
            for (int i = 0; i < crTables.Count; i++)
            {
                crTable = crTables[i];
                crTableLogOnInfo = crTable.LogOnInfo;
                crTableLogOnInfo.ConnectionInfo =
                crConnectionInfo;
                crTable.ApplyLogOnInfo(crTableLogOnInfo);

                //If your DatabaseName is changing at runtime, specify
                //the table location. For example, when you are reporting
                //off of a Northwind database on SQL server
                //you should have the following line of code:
                /*
                crTable.Location = "Northwind.dbo." +
                crTable.Location.Substring(crTable.Location.LastIndexOf(".") + 1)*/
            }


        }

        private void btnGetAgencyList_Click(object sender, EventArgs e)
        {
            try
            {
                parameters.Clear();
                parameters.Add("@AgencyID", 0);
                parameters.Add("@CityID", 0);
                //dsAgency = objDbAccess.FillData("usp_Get_Agency", parameters);
                dsAgency = objDbAccess.FillData("usp_Get_AgencyDistinct", parameters);
                lblAgencyCount.Text = lblAgencyCount.Text + " : " + dsAgency.Tables[0].Rows.Count.ToString();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnExportInExcel_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            DeviationReportView.Load(CurrentPath + "\\Reports\\LedgerReport.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
            ExportOptions CrExportOptions;

            var val = false;
            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();

                AgencyName = dr["Name"].ToString().Trim();

                if (AgencyName == "Brainchild Communications Pakistan (Pvt) Ltd" || AgencyName == "Pak Media Communications (Pvt.) Ltd." || AgencyName == "Z2C Pakistan (Pvt) Ltd" || AgencyName == "Blitz Advertising (Pvt.) Ltd. DDB" || AgencyName == "Fourays Advertising Pvt Ltd.")
                    val = true;
                if (val)
                {

                    DeviationReportView.SetParameterValue("@AgencyId", 0);
                    DeviationReportView.SetParameterValue("@BillingDate", dateTimePicker1.Value.ToShortDateString());
                    DeviationReportView.SetParameterValue("@AgingDays", 60);
                    DeviationReportView.SetParameterValue("@ChannelId", 0);
                    DeviationReportView.SetParameterValue("@ViewBy", 1);
                    DeviationReportView.SetParameterValue("@CityId", 0);
                    DeviationReportView.SetParameterValue("@Agency", AgencyName);



                    //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                    if (new FileInfo(CurrentPath + "\\LedgerBalance\\" + AgencyName + ".xls").Exists == true)
                    {
                        File.Delete(CurrentPath + "\\LedgerBalance\\" + AgencyName + ".xls");
                    }


                    CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\" + AgencyName + ".xls";
                    CrExportOptions = DeviationReportView.ExportOptions;
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                    DeviationReportView.Export();

                    Dispose();

                    GC.WaitForPendingFinalizers();

                    GC.Collect();

                }


                if (StopLoop == false)
                {
                    break;
                }

                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void btnExportInPDF_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            DeviationReportView.Load(CurrentPath + "\\Reports\\LedgerReport.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            PdfFormatOptions CrFormatTypeOptions = new PdfFormatOptions();
            ExportOptions CrExportOptions;



            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();

                AgencyName = dr["Name"].ToString();

                DeviationReportView.SetParameterValue("@AgencyId", 0);
                DeviationReportView.SetParameterValue("@BillingDate", dateTimePicker1.Value.ToShortDateString());
                DeviationReportView.SetParameterValue("@AgingDays", 60);
                DeviationReportView.SetParameterValue("@ChannelId", 0);
                DeviationReportView.SetParameterValue("@ViewBy", 1);
                DeviationReportView.SetParameterValue("@CityId", 0);
                DeviationReportView.SetParameterValue("@Agency", AgencyName);



                //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                if (new FileInfo(CurrentPath + "\\LedgerBalance\\" + AgencyName + ".pdf").Exists == true)
                {
                    File.Delete(CurrentPath + "\\LedgerBalance\\" + AgencyName + ".pdf");
                }


                CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\" + AgencyName + ".pdf";
                CrExportOptions = DeviationReportView.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                DeviationReportView.Export();



                if (StopLoop == false)
                {
                    break;
                }


                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            StopLoop = false;
        }

        private void btnOverDueInvoices_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            DeviationReportView.Load(CurrentPath + "\\Reports\\OverDueInvoices.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            string ClientNames = "";

            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
            ExportOptions CrExportOptions;



            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();

                // OverdueInvoicesReport.aspx?channelId=0&gency=4Square+Communications&%20AM&STI=0&criteria=0&orderby=1

                AgencyName = dr["Name"].ToString();

                DeviationReportView.SetParameterValue("@ChannelId", 0);
                DeviationReportView.SetParameterValue("@ViewBy", 1);
                DeviationReportView.SetParameterValue("@CityId", 0);
                DeviationReportView.SetParameterValue("@AgencyId", 0);
                DeviationReportView.SetParameterValue("@ClientId", 0);
                DeviationReportView.SetParameterValue("@clientIdstr", "");
                DeviationReportView.SetParameterValue("@Agency", AgencyName);
                //DeviationReportView.SetParameterValue("@Client", "");
                DeviationReportView.SetParameterValue("@Client", DBNull.Value);
                DeviationReportView.SetParameterValue("@clientNamestr", "");
                DeviationReportView.SetParameterValue("@FromDate", "1/1/1975");
                DeviationReportView.SetParameterValue("@ToDate", dateTimePicker1.Value.ToShortDateString());
                DeviationReportView.SetParameterValue("@DocumentId", 0);
                DeviationReportView.SetParameterValue("@Criteria", 0);
                DeviationReportView.SetParameterValue("@OrderBy", 1);

                //exec usp_GetOverDue_DeliveryDate @ChannelId=0,@ViewBy=1,@CityId=0,@AgencyId=0,@ClientId=0,@Agency=NULL,@Client=NULL,
                //  @FromDate='1999-03-01 00:00:00',@ToDate='2015-01-31 00:00:00',@DocumentId=0,@Criteria=0,@Orderby=1,@clientIdstr=N'',@clientNamestr=N''





                //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                if (new FileInfo(CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls").Exists == true)
                {
                    File.Delete(CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls");
                }


                CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls";
                CrExportOptions = DeviationReportView.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.Excel;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                DeviationReportView.Export();



                if (StopLoop == false)
                {
                    break;
                }

                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void btnStatementOfAccountEx_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();



            DeviationReportView.Load(CurrentPath + "\\Reports\\TransactionReport.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
            ExportOptions CrExportOptions;

            var cont = false;

            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();


                AgencyName = dr["Name"].ToString();
                //if (AgencyName == "Shazia Abbasi Consulting")
                //    cont = true;
                //if (cont == true)
                {

                    DeviationReportView.SetParameterValue("@AgencyId", 0);
                    DeviationReportView.SetParameterValue("@BillingDate", dateTimePicker1.Value.ToShortDateString());
                    DeviationReportView.SetParameterValue("@AgingDays", 60);
                    DeviationReportView.SetParameterValue("@ChannelId", 0);
                    DeviationReportView.SetParameterValue("@ViewBy", 1);
                    DeviationReportView.SetParameterValue("@CityId", 0);
                    DeviationReportView.SetParameterValue("@Agency", AgencyName);



                    //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                    if (new FileInfo(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls").Exists == true)
                    {
                        File.Delete(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls");
                    }


                    CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls";
                    CrExportOptions = DeviationReportView.ExportOptions;
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.Excel;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                    DeviationReportView.Export();



                    if (StopLoop == false)
                    {
                        break;
                    }
                }
                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void btnStatementOfAccountPDF_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            DeviationReportView.Load(CurrentPath + "\\Reports\\TransactionReport.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            PdfFormatOptions CrFormatTypeOptions = new PdfFormatOptions();
            ExportOptions CrExportOptions;



            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();
                var cont = false;
                AgencyName = dr["Name"].ToString();
                if (AgencyName == "Mindshare Pakistan (Private) Limited")
                    cont = true;
                if (cont == true)
                {

                    DeviationReportView.SetParameterValue("@AgencyId", 0);
                    DeviationReportView.SetParameterValue("@BillingDate", dateTimePicker1.Value.ToShortDateString());
                    DeviationReportView.SetParameterValue("@AgingDays", 60);
                    DeviationReportView.SetParameterValue("@ChannelId", 0);
                    DeviationReportView.SetParameterValue("@ViewBy", 1);
                    DeviationReportView.SetParameterValue("@CityId", 0);
                    DeviationReportView.SetParameterValue("@Agency", AgencyName);



                    //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                    if (new FileInfo(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".pdf").Exists == true)
                    {
                        File.Delete(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".pdf");
                    }


                    CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".pdf";
                    CrExportOptions = DeviationReportView.ExportOptions;
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                    DeviationReportView.Export();



                    if (StopLoop == false)
                    {
                        break;
                    }
                }

                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            DeviationReportView.Load(CurrentPath + "\\Reports\\OverDueInvoices.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            string ClientNames = "";

            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
            ExportOptions CrExportOptions;



            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();

                // OverdueInvoicesReport.aspx?channelId=0&gency=4Square+Communications&%20AM&STI=0&criteria=0&orderby=1

                AgencyName = dr["Name"].ToString();

                DeviationReportView.SetParameterValue("@ChannelId", 0);
                DeviationReportView.SetParameterValue("@ViewBy", 1);
                DeviationReportView.SetParameterValue("@CityId", 0);
                DeviationReportView.SetParameterValue("@AgencyId", 0);
                DeviationReportView.SetParameterValue("@ClientId", 0);
                DeviationReportView.SetParameterValue("@clientIdstr", "");
                DeviationReportView.SetParameterValue("@Agency", AgencyName);
                DeviationReportView.SetParameterValue("@Client", DBNull.Value);
                DeviationReportView.SetParameterValue("@clientNamestr", "");
                DeviationReportView.SetParameterValue("@FromDate", "1/1/1975");
                DeviationReportView.SetParameterValue("@ToDate", dateTimePicker1.Value.ToShortDateString());
                DeviationReportView.SetParameterValue("@DocumentId", 0);
                DeviationReportView.SetParameterValue("@Criteria", 0);
                DeviationReportView.SetParameterValue("@OrderBy", 1);







                //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                if (new FileInfo(CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls").Exists == true)
                {
                    File.Delete(CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls");
                }


                CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\OverDue_" + AgencyName + ".xls";
                CrExportOptions = DeviationReportView.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                DeviationReportView.Export();



                if (StopLoop == false)
                {
                    break;
                }

                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();

            DeviationReportView = new ReportDocument();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();



            DeviationReportView.Load(CurrentPath + "\\Reports\\TransactionReport.rpt");
            String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
            String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();

            DeviationReportView.SetDatabaseLogon(uid, pwd);

            ParameterFields paraFields = new ParameterFields();

            ApplyLoginInfo();

            string AgencyName = "";
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
            ExportOptions CrExportOptions;



            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                //DeviationReportView.ParameterFields.Clear();

                AgencyName = dr["Name"].ToString();

                DeviationReportView.SetParameterValue("@AgencyId", 0);
                DeviationReportView.SetParameterValue("@BillingDate", dateTimePicker1.Value.ToShortDateString());
                DeviationReportView.SetParameterValue("@AgingDays", 60);
                DeviationReportView.SetParameterValue("@ChannelId", 0);
                DeviationReportView.SetParameterValue("@ViewBy", 1);
                DeviationReportView.SetParameterValue("@CityId", 0);
                DeviationReportView.SetParameterValue("@Agency", AgencyName);



                //DeviationReportView.ExportToStream(ExportFormatType.PortableDocFormat);

                if (new FileInfo(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls").Exists == true)
                {
                    File.Delete(CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls");
                }


                CrDiskFileDestinationOptions.DiskFileName = CurrentPath + "\\LedgerBalance\\SOA_" + AgencyName + ".xls";
                CrExportOptions = DeviationReportView.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                DeviationReportView.Export();



                if (StopLoop == false)
                {
                    break;
                }

                //ReportDocument rpt = (CrystalDecisions.CrystalReports.Engine.ReportDocument)(Session["rpt_AttendanceRegisterDet"]);
                //DeviationReportView.Export(ExportFormatType.
                //DeviationReportView.ExportToHttpResponse(ExportFormatType.Excel, Response, true, Page.Title);
            }
            MessageBox.Show("Completed");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnNetBalance_Click(object sender, EventArgs e)
        {
            btnGetAgencyList.PerformClick();
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();
            DataTable dt = new DataTable();
            dt.Columns.Add("AgencyName");
            dt.Columns.Add("NetBalance");
            dt.Columns.Add("Year");

            var datatableBalances = new DataTable();
            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                var AgencyName = dr["Name"].ToString();

                if (AgencyName != "Brainchild Communications Pakistan (Pvt) Ltd" && AgencyName != "Group M Pakistan (Pvt.) Ltd." && AgencyName != "Blitz Advertising (Pvt.) Ltd. DDB" && AgencyName != "Creative Junction")
                {
                    clsDBAccess objAccess = new clsDBAccess();
                    datatableBalances = objAccess.GetNetBalances(0, dateTimePicker1.Value.ToShortDateString(), 60, 0, 1, 0, AgencyName);
                }
                //datatableBalances = datatableBalances.Sele
                if (datatableBalances.Rows.Count > 0)
                {

                    DataRow lastRow = datatableBalances.Rows[datatableBalances.Rows.Count - 1];
                    if (Convert.ToDateTime(lastRow["TransactionDate"]) > dateTimePicker1.Value.AddHours(23).AddMinutes(59))
                        lastRow = datatableBalances.Rows[datatableBalances.Rows.Count - 2];

                    DataRow row = dt.NewRow();
                    dt.Rows.Add(lastRow["AgencyName"].ToString(), lastRow["NetBalance"].ToString(), lastRow["TransactionDate"].ToString());

                }
            }

            var path = CurrentPath + "\\LedgerBalance\\AgencyLedgerBalance.csv";
            CreateCSVFile(dt, path);
        }

        public void CreateCSVFile(DataTable dt, string strFilePath)
        {
            try
            {
                // Create the CSV file to which grid data will be exported.
                StreamWriter sw = new StreamWriter(strFilePath, false);
                // First we will write the headers.
                //DataTable dt = m_dsProducts.Tables[0];
                int iColCount = dt.Columns.Count;
                for (int i = 0; i < iColCount; i++)
                {
                    sw.Write(dt.Columns[i]);
                    if (i < iColCount - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);

                // Now write all the rows.

                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 0; i < iColCount; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write(dr[i].ToString());
                        }
                        if (i < iColCount - 1)
                        {
                            sw.Write(",");
                        }
                    }

                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
