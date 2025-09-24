using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using ClosedXML.Excel;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;

namespace CrystalReportsApplication1
{

    public partial class GeneratePDFInvoice : Form
    {
        Hashtable parameters = new Hashtable();
        clsDBAccess objDbAccess = new clsDBAccess();
        public GeneratePDFInvoice()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            //label1.Text = "GST Calculation in Progress";
            //if (PrePareGSTData() == true)
            //{
            //    label1.Text = "GST Calculation completed, Proceeding to Export Data";
            //    ExportToExcel();
            //    label1.Text = "Client wise Aging with GST has been exported successfully";
            //}

            int ChannelId = 0;
            int Document = 1;

            DateTime TxDateFrom = dtFrom.Value;
            DateTime TxDateTo = dtTo.Value;
            

            DataTable dt = objDbAccess.GenerateInvoicePDF(TxDateFrom.Date.ToString("MM-dd-yyyy"), TxDateTo.Date.ToString("MM-dd-yyyy"), Document, "0", "0", ChannelId);

            if (dt != null)
            {
                SendPDFtoFolder(dt);

            }           
        }

        private string SendPDFtoFolder(DataTable dt)
        {
            int InvoiceID = 0;
            string InvoicePDF = "";
            string TCPDF = "";
            int i = 0;
            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    InvoiceID = Convert.ToInt32(dr[0]);
                    InvoicePDF = dr[1].ToString();
                    TCPDF = dr[2].ToString();
                    DataTable dtInvoice = new DataTable();
                    dtInvoice = objDbAccess.SendInvoicePDFParameters(InvoiceID);
                    PrintInvoice(dtInvoice, InvoiceID, InvoicePDF);
                    i++;
                }
            }
            catch (Exception ex)
            {
            }
            return "";
        }

        private bool PrintInvoice(DataTable dt, int InvoiceID, string TCPDF)
        {
            ReportDocument rpt = new ReportDocument();
            string InvoicePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            bool Result = false;
             string CurrentPath = System.IO.Directory.GetCurrentDirectory();

            //DeviationReportView.Load(CurrentPath + "\\Reports\\LedgerReport.rpt");


            //string path = Server.MapPath("~/Reports");
            rpt.Load(CurrentPath + "\\SalesTaxInvoice.rpt");
            rpt.SetDataSource(dt);
            rpt.SetParameterValue("@p_nSalesTax", 0);
            rpt.SetParameterValue("@p_nAgencyCommission", 0);
            rpt.SetParameterValue("@InvoiceId", InvoiceID);
            try
            {

                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                // string mpath = Server.MapPath("~/ExportPDF");
                CrDiskFileDestinationOptions.DiskFileName = InvoicePath +"\\Invoice\\" + TCPDF + ".pdf";
                CrExportOptions = rpt.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                rpt.Export();
                rpt.Dispose();
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
            finally
            {
                rpt.Close();
                rpt.Clone();
                rpt.Dispose();
                rpt = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return Result;
        }


        private void ExportToExcel()
        {
            DateTime AgingDate = dtTo.Value;
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();
            DataTable dt = objDbAccess.GetClientAgingMaster(AgingDate);
            int AgencyID = 0;
            int ClientID = 0;
            int Interval = 1;
            foreach (DataRow dr in dt.Rows)
            {
                AgencyID = Convert.ToInt32(dr[0]);
                ClientID = Convert.ToInt32(dr[1]);
                dr[2] = objDbAccess.GetAgencyName(Convert.ToInt32(dr[0]));
                dr[3] = objDbAccess.GetClientName(Convert.ToInt32(dr[1]));
                for (int i = 1; i < 10; i++)
                {
                    DataTable dtDetail = objDbAccess.GetClientAgingSlabs(AgencyID, ClientID, AgingDate, i);
                    DataTable dtGSTDetail = objDbAccess.GetClientAgingGSSlabs(AgencyID, ClientID, AgingDate, i);
                    if (dtDetail.Rows.Count > 0)
                    {
                        dr[i + 5] = Convert.ToInt32(dtDetail.Rows[0][0]);
                    }
                    else
                    {
                        dr[i + 5] = 0;
                    }

                    if (dtGSTDetail.Rows.Count > 0)
                    {
                        dr[i + 14] = Convert.ToInt32(dtGSTDetail.Rows[0][0]);
                    }
                    else
                    {
                        dr[i + 14] = 0; ;
                    }
                }
                Interval++;
            }
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "WorksheetName");
            string str = CurrentPath + "\\Aging\\" + "ClientWiseAging.xlsx";
            string TargetPath = str;
            wb.SaveAs(TargetPath);
        }

        private bool PrePareGSTData()
        {

            label1.Refresh();
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                parameters.Clear();
                parameters.Add("@AgingDate", dtTo.Value);
                objDbAccess.ExecuteNonQuery("Rpt_Aging_GST", parameters);


            }
            catch (Exception)
            {

            }
            return true;
        }

        private void ClientwiseAging_Load(object sender, EventArgs e)
        {

        }
    }
}
