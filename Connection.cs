using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Specialized;
//using System.Data.OracleClient;



    public class Connection
    {
        
        
        //NameValueCollection appSettings = ConfigurationManager.AppSettings; 
        private string Str_DBConString = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString(); 
        //private string Str_PrisonStringCon = "data source=RAFIQAWAN-PC; database=prison; User ID=sa; Password='NewPass102'";
        //private string strServerName = ConfigurationManager.ConnectionStrings["DBServerName"].ToString();
        private string strServerName = ConfigurationManager.AppSettings["DBServerName"].ToString();
        //private string strDatabaseName = ConfigurationManager.ConnectionStrings["DBName"].ToString();
        private string strDatabaseName = ConfigurationManager.AppSettings["DBName"].ToString();
        //private string strUserName = ConfigurationManager.ConnectionStrings["DBLoginInfo"].ToString();
        private string strUserName = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
        private string strPassword = ConfigurationManager.AppSettings["DBPwdInfo"].ToString(); 


        public Connection()
        {
          //  Str_PrisonStringCon = appSettings["PrisonConnectionString"].ToString();
            //for (int i = 0; i < appSettings.Count; i++)
            //{
                
            //     string str=i.ToString() +  appSettings.GetKey(i) + appSettings[i];
            //}

            //string str=ConfigurationSettings.AppSettings["PrisonConnectionString"].ToString();  
         
        }
        public string getstrConnection()
        {
            return Str_DBConString;
        }
        public SqlConnection getConnection()
        {
            return new SqlConnection(Str_DBConString);
        }

     

        public string getServer()
        {
            return strServerName;
        }
        public string getDatabase()
        {
            return (strDatabaseName);
        }
        public string getUserID()
        {
            return (strUserName);
        }
        public string getPassword()
        {
            return (strPassword);
        }
    }


