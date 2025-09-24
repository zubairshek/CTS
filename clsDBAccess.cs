using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Configuration;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows;
using CrystalDecisions.CrystalReports.Engine;
using System.Data.SqlClient;
using System.Collections;
using System.Windows;
using System.Windows.Forms;

namespace CrystalReportsApplication1
{
    public class clsDBAccess
    {
        string DBConStr = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
        SqlConnection conObj;// = new SqlConnection(DBConStr);

        public clsDBAccess()
        {
            conObj = new SqlConnection(DBConStr);
        }
        //public DataView Get_DV_Proc
        //         ObjDbAccess.BindCombo(ddlMS, "Name", "Id", "SP_GetMaritalStatus");

        //============= Calling Sample ========================
        //        Hashtable parameters = new Hashtable();
        //parameters.Add("@PrisonID", ddlPrison.SelectedValue);
        //parameters.Add("@Title", txtname.Text);
        //parameters.Add("@Date", cldDate.SelectedDate.ToShortDateString());
        //parameters.Add("@RosterType", ddlRosterType.SelectedValue);
        //parameters.Add("@RegisterID", HF_RegisterID.Value);
        ////parameters.Add("@RosterMasterID", HF_RosterMasterID.Value );
        //parameters.Add("@RegisterDesc", txtRegisterDesc.Text);
        //parameters.Add("@createdBy", user);
        ////parameters.Add("@createdOn", ddlRosterType.SelectedValue);                
        //parameters.Add("@IsActive", "1");

        //HF_ID.Value = objDB.InsertData("WarderDutyReg3121Master_insert", parameters);

        //============================================

        public string InsertData(string procedureName, Hashtable Parameters)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }

            SqlDataAdapter ad = new SqlDataAdapter();
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            string ID = "";
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ID = ds.Tables[0].Rows[0][0].ToString();
                }
            }
            //conObj.Open();         
            // ID = cmd.ExecuteScalar().ToString();
            //conObj.Close();
            return ID;
        }


        public DataView FillData(string procedureName, Hashtable Parameters, string orderBy)
        {
            DataView dv;
            SqlDataAdapter ad = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            dv = new DataView(ds.Tables[0]);
            dv.Sort = orderBy;

            return dv;
        }

        public DataSet FillData(string procedureName, Hashtable Parameters)
        {
            DataView dv;
            SqlDataAdapter ad = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {
                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            return ds;
        }
        public DataSet FillData(string procedureName)
        {
            DataView dv;
            SqlDataAdapter ad = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;

            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            return ds;
        }

        public bool ExecuteNonQuery(string procedureName, Hashtable Parameters)
        {
            bool Result = false;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            cmd.CommandTimeout = 6000;
            conObj.Open();
            try
            {
                ICollection ParamKeys = Parameters.Keys;
                foreach (object key in ParamKeys)
                {
                    cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
                }
                cmd.ExecuteNonQuery();
                Result = true;
            }
            catch (Exception)
            {

            }
            conObj.Close();
            return Result;
        }
        public DataView BindList(System.Windows.Forms.ComboBox lst, string ProcedureName, string orderBy, Hashtable Parameters, string TextField, string ValueField)
        {

            DataView dv;
            SqlDataAdapter ad = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = ProcedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            dv = new DataView(ds.Tables[0]);
            dv.Sort = orderBy;
            lst.DisplayMember = TextField;
            lst.ValueMember = ValueField;
            lst.DataSource = dv;

            //lst.DataTextField = TextField;
            //lst.DataValueField = ValueField;
            //lst.DataBind();
            return dv;
        }

        public DataView BindGridProc(System.Windows.Forms.DataGridView dg, string ProcedureName, string orderBy, Hashtable Parameters)
        {
            DataView dv;

            SqlDataAdapter ad = new SqlDataAdapter();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = ProcedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            dv = new DataView(ds.Tables[0]);
            if (ds.Tables.Count > 0)
            {
                dv.Sort = orderBy;
            }
            dg.DataSource = dv;
            //dg.DataBind();
            return dv;
        }

        public void BindCombo(System.Windows.Forms.ComboBox ddl, String Name, String ID, String ProcName)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conObj;
            cmd.CommandText = ProcName;
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter ad = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                ddl.DataSource = dt;
                ddl.DisplayMember = Name;
                ddl.ValueMember = ID;

            }
        }

        public void BindCheckListBox(System.Windows.Forms.CheckedListBox ChkLst, String Name, String ID, String ProcName)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conObj;
            cmd.CommandText = ProcName;
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter ad = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            ad.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                ChkLst.DataSource = dt;
                ChkLst.DisplayMember = Name;
                ChkLst.ValueMember = ID;

            }
        }

        public string DeleteData(string procedureName, Hashtable Parameters)
        {

            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }

            SqlDataAdapter ad = new SqlDataAdapter();
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);
            string ID = "";
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    ID = ds.Tables[0].Rows[0][0].ToString();
                }
            }
            //conObj.Open();         
            // ID = cmd.ExecuteScalar().ToString();
            //conObj.Close();
            return ID;
        }

        public DataSet SelectDataProc(string procedureName, Hashtable Parameters)
        {
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = procedureName;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conObj;
            //int loopCounter=0;
            ICollection ParamKeys = Parameters.Keys;
            foreach (object key in ParamKeys)
            {

                cmd.Parameters.Add(new SqlParameter(key.ToString(), Parameters[key.ToString()]));
            }

            SqlDataAdapter ad = new SqlDataAdapter();
            ad.SelectCommand = cmd;
            DataSet ds = new DataSet();
            ad.Fill(ds);

            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {

                }
            }

            return ds;

        }

        public DataTable GetClientAgingGSSlabs(int AgencyID, int ClientID, DateTime lastdayofmonth, int Interval)
        {

            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClientGSTSlabs", dbConn);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            dbAdapter.SelectCommand.Parameters.AddWithValue("@AgencyId", AgencyID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@ClientId", ClientID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@lastdayofmonth", lastdayofmonth);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@Interval", Interval);
            DataTable dt = new DataTable();
            dbAdapter.Fill(dt);
            dbConn.Close();
            return dt;
        }

        public DataTable GetClientAgingGSTFiscalYear(int AgencyID, int ClientID, DateTime from, DateTime to)
        {

            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClientGSTByFiscalYear", dbConn);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            dbAdapter.SelectCommand.Parameters.AddWithValue("@AgencyId", AgencyID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@ClientId", ClientID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@FromDate", from);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@TODate", to);
            DataTable dt = new DataTable();
            dbAdapter.Fill(dt);
            dbConn.Close();
            return dt;
        }

        public DataTable GetClientAgingSlabs(int AgencyID, int ClientID, DateTime lastdayofmonth, int Interval)
        {

            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClientAgingSlabs", dbConn);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            dbAdapter.SelectCommand.Parameters.AddWithValue("@AgencyId", AgencyID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@ClientId", ClientID);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@lastdayofmonth", lastdayofmonth);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@Interval", Interval);
            DataTable dt = new DataTable();
            dbAdapter.Fill(dt);
            dbConn.Close();
            return dt;
        }
        public DataTable GetClientAgingMaster(DateTime AgingDate)
        {
            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClientAgingMaster", dbConn);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@agingdate", AgingDate);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataTable dt = new DataTable();
            dbAdapter.Fill(dt);
            dbConn.Close();
            return dt;
        }

        public DataTable GetClientAgingMasterByFiscalYear(DateTime AgingDate)
        {
            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClientAgingMasterByFiscalYear", dbConn);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@agingdate", AgingDate);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataTable dt = new DataTable();
            dbAdapter.Fill(dt);
            dbConn.Close();
            return dt;
        }

        public DataTable SendInvoicePDFParameters(int InvoiceID)
        {
            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_SalesTaxInvoice", dbConn);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            dbAdapter.SelectCommand.CommandTimeout = 900;
            DataTable dtInvoices = new DataTable("Invoices");
            dbAdapter.SelectCommand.Parameters.AddWithValue("@p_nSalesTax", 0);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@p_nAgencyCommission", 0);
            dbAdapter.SelectCommand.Parameters.AddWithValue("@InvoiceId", InvoiceID);
            dbAdapter.Fill(dtInvoices);

            return dtInvoices;
        }


        public DataTable GenerateInvoicePDF(string dtFrom, string dtTo, int dateRange, string fromNo, string toNo, int channelId)
        {

            SqlConnection dbConn = new SqlConnection(DBConStr);
            SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_Get_Invoice_TCDatatableOnly", dbConn);
            dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;
            dbAdapter.SelectCommand.CommandTimeout = 1800;

            /*Input Parameters*/


            if (dtFrom != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@FromDate", dtFrom);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@FromDate", System.DBNull.Value);
            }

            if (dtTo != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ToDate", dtTo);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ToDate", System.DBNull.Value);
            }

            if (dateRange != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@DateRange", dateRange);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@DateRange", System.DBNull.Value);
            }

            if (fromNo != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@FromNo", fromNo);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@FromNo", System.DBNull.Value);
            }

            if (toNo != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ToNo", toNo);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ToNo", System.DBNull.Value);
            }

            if (channelId != null)
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ChannelValue", channelId);
            }
            else
            {
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ChannelValue", System.DBNull.Value);
            }




            DataTable dtInvoice = new DataTable("Invoice");

            dbAdapter.Fill(dtInvoice);

            return dtInvoice;
        }

        public string GetAgencyName(int MasterID)
        {

            try
            {
                SqlConnection dbConn = new SqlConnection(DBConStr);
                string strsql = " select Name  from tblAgencyGroup where ID =" + MasterID;
                SqlDataAdapter dbAdapter = new SqlDataAdapter(strsql, dbConn);
                dbAdapter.SelectCommand.CommandType = CommandType.Text;
                DataTable dt = new DataTable();
                dbAdapter.Fill(dt);
                dbConn.Close();
                return dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "NA";
            }

        }
        public string GetClientName(int ClientID)
        {

            try
            {
                SqlConnection dbConn = new SqlConnection(DBConStr);
                string strsql = " select Name from tblClient where ClientId= " + ClientID;
                SqlDataAdapter dbAdapter = new SqlDataAdapter(strsql, dbConn);
                dbAdapter.SelectCommand.CommandType = CommandType.Text;
                DataTable dt = new DataTable();
                dbAdapter.Fill(dt);
                dbConn.Close();
                return dt.Rows[0][0].ToString();
            }
            catch (Exception)
            {
                return "NA";
            }
        }

        public DataTable GetNetBalances(int agencyId, string billingDate, int agingDays, int channelId, int viewBy, int cityId, string agencyName)
        {
            SqlConnection dbConn = new SqlConnection(DBConStr);
            try
            {
                SqlDataAdapter dbAdapter = new SqlDataAdapter("usp_GetClosingLedgerBalance", dbConn);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@AgencyId", agencyId);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@BillingDate", billingDate);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@AgingDays", agingDays);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ChannelId", channelId);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@ViewBy", viewBy);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@CityId", cityId);
                dbAdapter.SelectCommand.Parameters.AddWithValue("@Agency", agencyName);
                dbAdapter.SelectCommand.CommandType = CommandType.StoredProcedure;

                DataTable dt = new DataTable();
                dbAdapter.Fill(dt);
                dbConn.Close();
                return dt;
            }
            catch
            {
                return new DataTable();
            }
            finally
            {
                dbConn.Close();
            }
        }
    }
}
