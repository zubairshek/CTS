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

namespace CrystalReportsApplication1
{

    public partial class ClientwiseAging : Form
    {
        Hashtable parameters = new Hashtable();
        clsDBAccess objDbAccess = new clsDBAccess();
        public ClientwiseAging()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "GST Calculation in Progress";
            if (PrePareGSTData() == true)
            {
                label1.Text = "GST Calculation completed, Proceeding to Export Data";
                ExportToExcel();                
                label1.Text = "Client wise Aging with GST has been exported successfully";
            }
        }

        private void ExportToExcel()
        {
            DateTime AgingDate = dateTimePicker1.Value;
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

        private void ExportGSTToExcelFiscalYear()
        {
            DateTime AgingDate = dateTimePicker1.Value;
            string CurrentPath = System.IO.Directory.GetCurrentDirectory();
            DataTable dt = objDbAccess.GetClientAgingMasterByFiscalYear(AgingDate);
            int AgencyID = 0;
            int ClientID = 0;
           
            foreach (DataRow dr in dt.Rows)
            {
                int Interval = 1;
                AgencyID = Convert.ToInt32(dr[0]);
                ClientID = Convert.ToInt32(dr[1]);
                dr[2] = objDbAccess.GetAgencyName(Convert.ToInt32(dr[0]));
                dr[3] = objDbAccess.GetClientName(Convert.ToInt32(dr[1]));
                for (int i = 2008; i <= 2022; i++)
                {
                    var fromDate = new DateTime(i - 1, 7, 1);
                    var toDate = new DateTime(i, 6, 30);

                    DataTable dtGSTDetail = objDbAccess.GetClientAgingGSTFiscalYear(AgencyID, ClientID, fromDate, toDate);

                    if (dtGSTDetail.Rows.Count > 0)
                    {
                        dr[Interval + 5] = Convert.ToInt32(dtGSTDetail.Rows[0][0]);
                    }
                    else
                    {
                        dr[Interval + 5] = 0; ;
                    }
                    Interval++;
                }
               
            }
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "WorksheetName");
            string str = CurrentPath + "\\Aging\\" + "Aging-ClientFiscalYearWise.xlsx";
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
                parameters.Add("@AgingDate", dateTimePicker1.Value);
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

        private void button2_Click(object sender, EventArgs e)
        {
            label1.Text = "GST Calculation in Progress";
            if (PrePareGSTData() == true)
            {
                label1.Text = "GST Calculation completed, Proceeding to Export Data";
                ExportGSTToExcelFiscalYear();
                label1.Text = "Client wise Aging with GST has been exported successfully";
            }
        }
    }
}
