using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CrystalReportsApplication1
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //SOA
            //Application.Run(new OverDueAgencyWise());
            //Application.Run(new ClientwiseAging());
            //LedgerBalance
            //Application.Run(new Form1());
            //Application.Run(new GeneratePDFInvoice());
            //Application.Run(new frmYearwise());
            //DailyRevenue
            Application.Run(new DailyRevenueReport());

        }
    }
}
