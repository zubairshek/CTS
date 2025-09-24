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
using ClosedXML.Excel;

namespace CrystalReportsApplication1
{
    public partial class OverDueAgencyWise : Form
    {
        Hashtable parameters = new Hashtable();
        clsDBAccess objDbAccess = new clsDBAccess();
        public OverDueAgencyWise()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PopulateReleaseOrderView(0);
        }

        protected void PopulateReleaseOrderView(int action)
        {
            label1.Text = "Exporting in Progress";
            label1.Refresh();
            Cursor.Current = Cursors.WaitCursor;
            if (action == 0)
            {
                var channelId = 1;
                parameters.Clear();
                parameters.Add("@AgencyID", 0);
                parameters.Add("@CityID", 0);
               
                //DataSet  dsAgency = objDbAccess.FillData("usp_Get_Agency", parameters);
                parameters.Add("@ChannelId", channelId);
                DataSet dsAgency = objDbAccess.FillData("usp_Get_AgencyByChannel", parameters);


                progressBar1.Minimum = 0;
                progressBar1.Maximum = dsAgency.Tables[0].Rows.Count;
                int i = 0;
                foreach (DataRow dr in dsAgency.Tables[0].Rows)
                {
                    parameters.Clear();
                    //parameters.Add("@ChannelId", 0);
                    parameters.Add("@ChannelId", channelId);
                    parameters.Add("@ViewBy", 1);
                    parameters.Add("@CityId", 0);
                    parameters.Add("@AgencyId",0);
                    parameters.Add("@ClientId", 0);
                    parameters.Add("@Agency",dr[1].ToString ()) ;
                    parameters.Add("@Client",DBNull.Value );
                    parameters.Add("@FromDate","2008-01-01");
                    parameters.Add("@ToDate",dateTimePicker1.Value);
                    parameters.Add("@DocumentId", 0);
                    parameters.Add("@Criteria", 0);
                    parameters.Add("@Orderby", 1);
                    parameters.Add("@clientIdstr", " ");
                    parameters.Add("@clientNamestr"," ");   
            

                    DataSet dsAgencyOverDue = objDbAccess.FillData("usp_GetOverDue", parameters);
                    CreateExcel (dsAgencyOverDue.Tables[0]); 
                    i++;
                    progressBar1.Value = i;
                }
                label1.Text = "Export Completed.....";
                Cursor.Current = Cursors.Default;
               // objInvoiceDB = new CTS.InvoiceDB();
                //DataTable dtInvoices = objInvoiceDB.GetOverDueInvoices(cityId, agencyId, clientId, orderby, channelId, FromDatePicker.SelectedDate, ToDatePicker.SelectedDate, true, int.Parse(ddlSTI.SelectedItem.Value));
                // by Aijaz - 01 Feb 2013
                //DataTable dtInvoices = objInvoiceDB.GetOverDue(channelId, viewby, cityId, agencyId, clientId, agency, client, FromDatePicker.SelectedDate, ToDatePicker.SelectedDate, documentId, criteria, orderby, clientIdstr, clientNamestr);
                //DataTable dtInvoices = objInvoiceDB.GetOverDue_DeliveryDate(channelId, viewby, cityId, agencyId, clientId, agency, client, FromDatePicker.SelectedDate, ToDatePicker.SelectedDate, documentId, criteria, orderby, clientIdstr, clientNamestr, IsGovt);
               // CreateExcel(dtInvoices);
                //  gvOutstandingInvoice.DataSource = dtInvoices;
                // gvOutstandingInvoice.DataBind();
            }
           

        }
        private void CreateExcel(DataTable dt)
        {

            try
            {
                XLWorkbook wb = new XLWorkbook();
                var ws = wb.AddWorksheet("Over Due Summary");

                ws.Cell("B2").Value = "Overdue Invoices as On " + dateTimePicker1.Value.ToString ("dd/MM/yyyy");
                ws.Cell("B2").Style.Font.SetFontSize(14);
                ws.Range("B2:D2").Style.Fill.BackgroundColor = XLColor.CornflowerBlue;                
                 ws.Range("B2:D2").Merge();
                //rngTable.Cell(1, 1).Style.Font.Bold = true;
                //rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                //rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                //rngTable.Row(1).Merge();

                ws.Cell("A6").Value = "Sr.#";
                ws.Cell("B6").Value = "Channel";
                ws.Cell("C6").Value = "City";
                ws.Cell("D6").Value = "Agency";
                ws.Cell("E6").Value = "Client";
                ws.Cell("F6").Value = "GST #";
                ws.Cell("G6").Value = "STI #";
                ws.Cell("H6").Value = "Invoice #";
                ws.Cell("I6").Value = "Invoice Date";
                ws.Cell("J6").Value = "Reference";
                ws.Cell("K6").Value = "Total Amount";
                ws.Cell("L6").Value = "Status";
                ws.Cell("M6").Value = "Rem. Amount";
                ws.Cell("N6").Value = "G.S.Tax";
                ws.Cell("O6").Value = "Net Amount";
                ws.Cell("P6").Value = "Adj.Reason";
                ws.Cell("Q6").Value = "Actual GST";
                ws.Cell("R6").Value = "Un-InvoiceDate";
                ws.Row(6).Cells("1:25").Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                //row1.Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                ws.Column(2).Width = 20;
                ws.Column(3).Width = 20;
                ws.Column(4).Width = 30;
                ws.Column(5).Width = 35;
                ws.Column(6).Width = 15;
                ws.Column(7).Width = 15;
                ws.Column(8).Width = 15;
                ws.Column(9).Width = 15;
                ws.Column(10).Width = 15;
                ws.Column(11).Width = 14;
                ws.Column(12).Width = 14;
                ws.Column(13).Width = 14;
                ws.Column(14).Width = 14;
                ws.Column(15).Width = 14;
                ws.Column(16).Width = 14;
                ws.Column(17).Width = 14;
                ws.Column(18).Width = 16;
                

                //ws.Columns().AdjustToContents();
                //rngTable.Cell(1, 1).Style.Font.Bold = true;
                //rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
                //rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                //rngTable.Row(1).Merge();
                int i = 7;
                string AgencyName = "";
                foreach (DataRow row in dt.Rows)
                {                   
                
                        AgencyName = row[13].ToString();
                        ws.Cell("A" + i.ToString()).Value = (i - 6).ToString().Trim();
                        ws.Cell("B" + i.ToString()).Value = row[20].ToString().Trim();
                        ws.Cell("C" + i.ToString()).Value = row[10].ToString().Trim();
                        ws.Cell("D" + i.ToString()).Value = row[13].ToString().Trim();
                        ws.Cell("E" + i.ToString()).Value = row[15].ToString().Trim();

                        //ws.Cell("F" + i.ToString()).CellLeft();//= . = true;//.al.DataType = XLCellValues.Number;
                        ws.Cell("F" + i.ToString()).Value = row[16].ToString().Trim();
                        try
                        {
                            ws.Cell("F" + i.ToString()).Style.NumberFormat.Format = "#,##0";
                            ws.Cell("F" + i.ToString()).DataType = XLCellValues.Text;
                        }
                        catch (Exception)
                        {
                            //   throw;
                        }


                        ws.Cell("G" + i.ToString()).Value = row[25].ToString().Trim();
                        ws.Cell("H" + i.ToString()).Value = row[1].ToString().Trim();
                        try
                        {
                            ws.Cell("I" + i.ToString()).Value = Convert.ToDateTime(row[2]).ToString("dd/MM/yyyy");// .ToString().Trim();

                        }
                        catch (Exception)
                        {

                        }

                        ws.Cell("J" + i.ToString()).Value = row[8].ToString().Trim();
                        ws.Cell("K" + i.ToString()).Value = row[3].ToString().Trim();
                        ws.Cell("L" + i.ToString()).Value = row[5].ToString().Trim();
                        ws.Cell("M" + i.ToString()).Value = row[4].ToString().Trim();
                        ws.Cell("N" + i.ToString()).Value = row[23].ToString().Trim();
                        ws.Cell("O" + i.ToString()).Value = row[29].ToString().Trim();
                        ws.Cell("P" + i.ToString()).Value = row[30].ToString().Trim();
                        ws.Cell("Q" + i.ToString()).Value = row[22].ToString().Trim();
                        ws.Cell("R" + i.ToString()).Value = Convert.ToDateTime(row[2]);
                      
                        i++;
                    }
                  
                
                string TargetPath = "E:\\AgingFiles\\" + AgencyName + ".xlsx";
                wb.SaveAs(TargetPath);
            }
            catch (Exception ex)
            {
                //lblMsg2.Text = ex.Message;

                //  throw;
            }




        }
    }
}
