using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.Collections;
using CrystalDecisions.Windows.Forms;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;
namespace CrystalReportsApplication1
{
    public partial class frmYearwise : Form
    {
        public frmYearwise()
        {
            InitializeComponent();
        }
        clsDBAccess objDbAccess = new clsDBAccess();
        Connection objCon = new Connection();
        Hashtable parameters = new Hashtable();
        String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
        String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();
        DataSet dsAgency = new DataSet();
        DataSet dsYear = new System.Data.DataSet();
        private void frmYearwise_Load(object sender, EventArgs e)
        {



            parameters.Clear();
            parameters.Add("@ID", 0);           
            dsAgency = objDbAccess.FillData("usp_Get_AgencyGroup", parameters);


        }

        private void button1_Click(object sender, EventArgs e)
        {
              string CurrentPath = System.IO.Directory.GetCurrentDirectory();
            parameters.Clear();
            parameters.Add("@AgencyID", 0);
            dsYear = objDbAccess.FillData("usp_GetYearList", parameters);
            DataSet _ds = new System.Data.DataSet();
            DataTable MasterTable = new System.Data.DataTable();
            MasterTable.Columns.Add("AgencyID", typeof(Int32));
            MasterTable.Columns.Add("AgencyName", typeof(string));
            foreach (DataRow dryear in dsYear.Tables[0].Rows)
            {
                try
                {
                    MasterTable.Columns.Add(Convert.ToString(dryear[0]),typeof(Decimal));
                }
                catch (Exception ex)
                {
                }
            }
            int count = MasterTable.Columns.Count ;
            DataRow row = MasterTable.NewRow();
            foreach (DataRow dr in dsAgency.Tables[0].Rows)
            {
                
                int i = 0;
                row = MasterTable.NewRow();
                row[i] = Convert.ToInt32(dr["ID"]);
                row[i + 1] = dr["Name"].ToString();
                i=2;
                foreach (DataRow dryear in dsYear.Tables[0].Rows)
                {
                
                    try
                    {
                       
                            parameters.Clear();
                            parameters.Add("@Year", Convert.ToInt32(dryear[0]));
                            parameters.Add("@AgencyId", Convert.ToInt32(dr["ID"]));
                            _ds = objDbAccess.FillData("usp_GetYearTotal", parameters);
                            row[i] = Convert.ToDecimal(_ds.Tables[0].Rows[0][0]);
                        
                         
                        i++;
                    }
                    catch (Exception)
                    {

                       
                    }

                }

                MasterTable.Rows.Add(row);
              
            }
            XLWorkbook wb = new XLWorkbook();
            DataTable dt = MasterTable;
            wb.Worksheets.Add(dt, "WorksheetName");
            string str = CurrentPath + "\\LedgerBalance\\" + "Munir.xlsx";
            string  TargetPath = str ;
            wb.SaveAs(TargetPath);
           
           
           
            
            

            
         //   MasterTable.WriteXml(str); 
        }
    }
}
