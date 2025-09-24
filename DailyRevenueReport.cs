using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using System.Configuration;
using CrystalDecisions.Shared;
using System.Collections;
using System.IO;
using System.Net.Mail;
using System.Net;
using iTextSharp.text.pdf;
using CrystalDecisions.Windows.Forms;
using System.Drawing.Imaging;
using ImageMagick;

namespace CrystalReportsApplication1
{
    public partial class DailyRevenueReport : Form
    {
        ReportDocument DeviationReportView = new ReportDocument();
        String uid = ConfigurationManager.AppSettings["DBLoginInfo"].ToString();
        String pwd = ConfigurationManager.AppSettings["DBPwdInfo"].ToString();
        clsDBAccess objDbAccess = new clsDBAccess();
        Connection objCon = new Connection();
        System.Collections.Hashtable parameters = new System.Collections.Hashtable();
        public DailyRevenueReport()
        {
            InitializeComponent();
            this.Load += DailyRevenueReport_Load;
        }
        private void DailyRevenueReport_Load(object sender, EventArgs e)
        {
            GenerateReport();
            //GenerateMonthlyRevenueReport();
        }

        private DataSet GetData(int channelId, DateTime fromDate, DateTime toDate)
        {
            parameters.Clear();
            parameters.Add("@ChannelId", channelId);
            parameters.Add("@FromDate", fromDate);
            parameters.Add("@ToDate", toDate);
            var ds = objDbAccess.FillData("usp_GetDailyRevenue", parameters);
            return ds;
        }

        private DataSet GetMonthlyRevenueData(int channelId, DateTime fromDate, DateTime toDate)
        {
            parameters.Clear();
            parameters.Add("@ChannelId", channelId);
            parameters.Add("@FromDate", fromDate);
            parameters.Add("@ToDate", toDate);
            var ds = objDbAccess.FillData("usp_GetDailyRevenue", parameters);
            return ds;
        }

        private void GenerateMonthlyRevenueReport()
        {
            List<int> lstChannels = new List<int> { 0 };
            MailMessage mail = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            foreach (var channelId in lstChannels)
            {
                DeviationReportView = new ReportDocument();
                string reportName = "";
                if (channelId == 0)
                    reportName = "MonthlyRevenueReport.rpt";

                string reportPath = Path.Combine(Application.StartupPath, "Reports", reportName);
                DeviationReportView.Load(reportPath, OpenReportMethod.OpenReportByTempCopy);
                string uid = ConfigurationManager.AppSettings["DBLoginInfo"];
                string pwd = ConfigurationManager.AppSettings["DBPwdInfo"];
                DeviationReportView.SetDatabaseLogon(uid, pwd);
                ApplyLoginInfo();
                //DateTime currentDate = DateTime.Now;
                //DateTime firstDateOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);

                DateTime now = DateTime.Now.AddDays(-1);
                // First day of the current month
                DateTime firstDay = new DateTime(now.Year, now.Month, 1);
                // Last day of the current month
                DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

                var ds = GetMonthlyRevenueData(channelId, firstDay, lastDay);
                DeviationReportView.SetParameterValue("@ChannelId", channelId);
                DeviationReportView.SetParameterValue("@FromDate", firstDay);
                DeviationReportView.SetParameterValue("@ToDate", lastDay);
                var fileName = "MonthlyRevenueReport.pdf";
                string pdfPath = Path.Combine(Path.GetTempPath(), $"{fileName}");
                DeviationReportView.ExportToDisk(ExportFormatType.PortableDocFormat, pdfPath);
                List<string> imagePaths = new List<string>();
                using (var images = new MagickImageCollection())
                {
                    MagickReadSettings settings = new MagickReadSettings
                    {
                        Density = new Density(200)
                    };
                    images.Read(pdfPath, settings);
                    int page = 1;
                    foreach (var image in images)
                    {
                        image.Format = MagickFormat.Jpg;
                        image.Quality = 90;
                        fileName = "MonthlyRevenueReport.jpg";
                        string imagePath = Path.Combine(Path.GetTempPath(), $"{fileName}");
                        image.Write(imagePath);
                        imagePaths.Add(imagePath);
                        page++;
                    }
                }
                foreach (string imagePath in imagePaths)
                {
                    mail.Attachments.Add(new Attachment(imagePath));
                }
            }
            smtp = CreateEmailMonthlyRevenue(mail);
            smtp.Send(mail);
            Environment.Exit(0);
        }

        private void GenerateReport()
        {
            List<int> lstChannels = new List<int> { 0, 1, 3 };
            MailMessage mail = new MailMessage();
            SmtpClient smtp = new SmtpClient();
            foreach (var channelId in lstChannels)
            {
                DeviationReportView = new ReportDocument();
                string reportName = "";
                if (channelId == 0)
                    reportName = "DRRC.rpt";
                else if (channelId == 1)
                    reportName = "DRRNews.rpt";
                else
                    reportName = "DRR.rpt";
                string reportPath = Path.Combine(Application.StartupPath, "Reports", reportName);
                DeviationReportView.Load(reportPath, OpenReportMethod.OpenReportByTempCopy);
                string uid = ConfigurationManager.AppSettings["DBLoginInfo"];
                string pwd = ConfigurationManager.AppSettings["DBPwdInfo"];
                DeviationReportView.SetDatabaseLogon(uid, pwd);
                ApplyLoginInfo();
                DateTime currentDate = DateTime.Now.AddDays(-1);
                DateTime firstDateOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
                var ds = GetData(channelId, firstDateOfMonth, currentDate);

                //// First day of the current month
                //DateTime firstDay = new DateTime(DateTime.Now.Year, 6, 1);
                //// Last day of the current month
                //DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);
                //var ds = GetData(channelId, firstDay, lastDay);

                DeviationReportView.SetParameterValue("@ChannelId", channelId);
                DeviationReportView.SetParameterValue("@FromDate", firstDateOfMonth);
                DeviationReportView.SetParameterValue("@ToDate", currentDate);
                var fileName = channelId == 0 ? "DailyRevenue_ET_EN_Combined.pdf" : channelId == 1 ? "DailyRevenue_EN.pdf" : channelId == 3 ? "DailyRevenue_ET.pdf" : "notfound.pdf";
                string pdfPath = Path.Combine(Path.GetTempPath(), $"{fileName}");

                DeviationReportView.ExportToDisk(ExportFormatType.PortableDocFormat, pdfPath);

                List<string> imagePaths = new List<string>();
                using (var images = new MagickImageCollection())
                {
                    MagickReadSettings settings = new MagickReadSettings
                    {
                        Density = new Density(200)
                    };
                    images.Read(pdfPath, settings);
                    int page = 1;
                    foreach (var image in images)
                    {
                        image.Format = MagickFormat.Jpg;
                        image.Quality = 90;
                        fileName = channelId == 0 ? "DailyRevenue_ET_EN_Combined.jpg" : channelId == 1 ? "DailyRevenue_EN.jpg" : channelId == 3 ? "DailyRevenue_ET.jpg" : "notfound.jpg";
                        string imagePath = Path.Combine(Path.GetTempPath(), $"{fileName}");
                        image.Write(imagePath);
                        imagePaths.Add(imagePath);
                        page++;
                    }
                }
                foreach (string imagePath in imagePaths)
                {
                    mail.Attachments.Add(new Attachment(imagePath));
                }
            }
            smtp = CreateEmail(mail);
            smtp.Send(mail);
            Environment.Exit(0);
        }
        protected void ApplyLoginInfoForReport(ReportDocument reportDoc)
        {
            ConnectionInfo crConnectionInfo = new ConnectionInfo();
            crConnectionInfo.ServerName = ConfigurationManager.AppSettings["DBServerName"];
            crConnectionInfo.UserID = ConfigurationManager.AppSettings["DBLoginInfo"];
            crConnectionInfo.Password = ConfigurationManager.AppSettings["DBPwdInfo"];
            Database crDatabase = reportDoc.Database;
            Tables crTables = crDatabase.Tables;
            CrystalDecisions.CrystalReports.Engine.Table crTable;
            TableLogOnInfo crTableLogOnInfo;
            for (int i = 0; i < crTables.Count; i++)
            {
                crTable = crTables[i];
                crTableLogOnInfo = crTable.LogOnInfo;
                crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
                crTable.ApplyLogOnInfo(crTableLogOnInfo);
            }
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
            for (int i = 0; i < crTables.Count; i++)
            {
                crTable = crTables[i];
                crTableLogOnInfo = crTable.LogOnInfo;
                crTableLogOnInfo.ConnectionInfo = crConnectionInfo;
                crTable.ApplyLogOnInfo(crTableLogOnInfo);
            }
        }
        private static SmtpClient CreateEmail(MailMessage mail)
        {
            mail.From = new MailAddress("mmsadmin@expressnews.tv", "CTS - Daily Revenue Report");
            mail.To.Add("muhammad.zubair@express.com.pk");
            //mail.CC.Add("munir.mustafa@express.com.pk");
            //mail.CC.Add("Mohtarim.ali@expressnews.tv");
            //mail.CC.Add("mohammad.ehsan@expressnews.tv");
            //mail.CC.Add("razi.hasan@expressnews.tv");
            mail.Subject = "JPG format - Monthly Revenue Report";
            mail.Body = "Please find the attached CTS Daily Revenue Report in JPG format.";
            SmtpClient smtp = new SmtpClient("175.107.197.118")
            {
                Port = 587,
                Credentials = new NetworkCredential("mmsadmin@expressnews.tv", "W@ll3tb0x"),
                EnableSsl = false
            };
            return smtp;
        }

        private static SmtpClient CreateEmailMonthlyRevenue(MailMessage mail)
        {
            mail.From = new MailAddress("mmsadmin@expressnews.tv", "CTS - Monthly Revenue Report");
            mail.To.Add("muhammad.zubair@express.com.pk");
            //mail.CC.Add("munir.mustafa@express.com.pk");
            //mail.CC.Add("Mohtarim.ali@expressnews.tv");
            //mail.CC.Add("mohammad.ehsan@expressnews.tv");
            mail.Subject = "JPG format - Monthly Revenue Report";
            mail.Body = "Please find the attached CTS Monthly Revenue Report in JPG format.";
            SmtpClient smtp = new SmtpClient("175.107.197.118")
            {
                Port = 587,
                Credentials = new NetworkCredential("mmsadmin@expressnews.tv", "W@ll3tb0x"),
                EnableSsl = false
            };
            return smtp;
        }
    }
}