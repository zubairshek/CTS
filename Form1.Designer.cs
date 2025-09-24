namespace CrystalReportsApplication1
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGetAgencyList = new System.Windows.Forms.Button();
            this.btnExportInExcel = new System.Windows.Forms.Button();
            this.btnExportInPDF = new System.Windows.Forms.Button();
            this.lblAgencyCount = new System.Windows.Forms.Label();
            this.btnStop = new System.Windows.Forms.Button();
            this.btnOverDueInvoices = new System.Windows.Forms.Button();
            this.btnStatementOfAccountEx = new System.Windows.Forms.Button();
            this.btnStatementOfAccountPDF = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.CrystalReport11 = new CrystalReportsApplication1.CrystalReport1();
            this.btnNetBalance = new System.Windows.Forms.Button();
            this.cachedSalesTaxInvoice1 = new CrystalReportsApplication1.CachedSalesTaxInvoice();
            this.cachedSalesTaxInvoice2 = new CrystalReportsApplication1.CachedSalesTaxInvoice();
            this.cachedSalesTaxInvoice3 = new CrystalReportsApplication1.CachedSalesTaxInvoice();
            this.cachedSalesTaxInvoice4 = new CrystalReportsApplication1.CachedSalesTaxInvoice();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(104, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ledger Detail till date";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(236, 45);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Location = new System.Drawing.Point(35, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(185, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "Export Ledger Report for All agencies";
            // 
            // btnGetAgencyList
            // 
            this.btnGetAgencyList.Location = new System.Drawing.Point(665, 15);
            this.btnGetAgencyList.Name = "btnGetAgencyList";
            this.btnGetAgencyList.Size = new System.Drawing.Size(122, 23);
            this.btnGetAgencyList.TabIndex = 3;
            this.btnGetAgencyList.Text = "(auto)Get Agency List";
            this.btnGetAgencyList.UseVisualStyleBackColor = true;
            this.btnGetAgencyList.Click += new System.EventHandler(this.btnGetAgencyList_Click);
            // 
            // btnExportInExcel
            // 
            this.btnExportInExcel.Location = new System.Drawing.Point(35, 122);
            this.btnExportInExcel.Name = "btnExportInExcel";
            this.btnExportInExcel.Size = new System.Drawing.Size(185, 23);
            this.btnExportInExcel.TabIndex = 4;
            this.btnExportInExcel.Text = "Agency Ledger Balance in Excel";
            this.btnExportInExcel.UseVisualStyleBackColor = true;
            this.btnExportInExcel.Click += new System.EventHandler(this.btnExportInExcel_Click);
            // 
            // btnExportInPDF
            // 
            this.btnExportInPDF.Location = new System.Drawing.Point(228, 122);
            this.btnExportInPDF.Name = "btnExportInPDF";
            this.btnExportInPDF.Size = new System.Drawing.Size(165, 23);
            this.btnExportInPDF.TabIndex = 5;
            this.btnExportInPDF.Text = "Ledger Balance in PDF";
            this.btnExportInPDF.UseVisualStyleBackColor = true;
            this.btnExportInPDF.Click += new System.EventHandler(this.btnExportInPDF_Click);
            // 
            // lblAgencyCount
            // 
            this.lblAgencyCount.AutoSize = true;
            this.lblAgencyCount.Location = new System.Drawing.Point(32, 93);
            this.lblAgencyCount.Name = "lblAgencyCount";
            this.lblAgencyCount.Size = new System.Drawing.Size(156, 13);
            this.lblAgencyCount.TabIndex = 6;
            this.lblAgencyCount.Text = "Agencies Available for Report : ";
            // 
            // btnStop
            // 
            this.btnStop.Location = new System.Drawing.Point(399, 122);
            this.btnStop.Name = "btnStop";
            this.btnStop.Size = new System.Drawing.Size(122, 23);
            this.btnStop.TabIndex = 8;
            this.btnStop.Text = "Stop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // btnOverDueInvoices
            // 
            this.btnOverDueInvoices.Location = new System.Drawing.Point(35, 151);
            this.btnOverDueInvoices.Name = "btnOverDueInvoices";
            this.btnOverDueInvoices.Size = new System.Drawing.Size(223, 23);
            this.btnOverDueInvoices.TabIndex = 9;
            this.btnOverDueInvoices.Text = "Export Overdue Report in Excel";
            this.btnOverDueInvoices.UseVisualStyleBackColor = true;
            this.btnOverDueInvoices.Click += new System.EventHandler(this.btnOverDueInvoices_Click);
            // 
            // btnStatementOfAccountEx
            // 
            this.btnStatementOfAccountEx.Location = new System.Drawing.Point(35, 180);
            this.btnStatementOfAccountEx.Name = "btnStatementOfAccountEx";
            this.btnStatementOfAccountEx.Size = new System.Drawing.Size(223, 23);
            this.btnStatementOfAccountEx.TabIndex = 10;
            this.btnStatementOfAccountEx.Text = "Export Statement of Account in Excel";
            this.btnStatementOfAccountEx.UseVisualStyleBackColor = true;
            this.btnStatementOfAccountEx.Click += new System.EventHandler(this.btnStatementOfAccountEx_Click);
            // 
            // btnStatementOfAccountPDF
            // 
            this.btnStatementOfAccountPDF.Location = new System.Drawing.Point(264, 180);
            this.btnStatementOfAccountPDF.Name = "btnStatementOfAccountPDF";
            this.btnStatementOfAccountPDF.Size = new System.Drawing.Size(257, 23);
            this.btnStatementOfAccountPDF.TabIndex = 11;
            this.btnStatementOfAccountPDF.Text = "Export Statement of Account in PDF";
            this.btnStatementOfAccountPDF.UseVisualStyleBackColor = true;
            this.btnStatementOfAccountPDF.Click += new System.EventHandler(this.btnStatementOfAccountPDF_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(264, 151);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(257, 23);
            this.button1.TabIndex = 12;
            this.button1.Text = "Export Overdue Report in Excel data only";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(35, 209);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(486, 23);
            this.button2.TabIndex = 13;
            this.button2.Text = "Export Statement of Account in Excel Data only format";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btnNetBalance
            // 
            this.btnNetBalance.Location = new System.Drawing.Point(35, 238);
            this.btnNetBalance.Name = "btnNetBalance";
            this.btnNetBalance.Size = new System.Drawing.Size(223, 23);
            this.btnNetBalance.TabIndex = 14;
            this.btnNetBalance.Text = "Agency Ledger Net Balance in Excel";
            this.btnNetBalance.UseVisualStyleBackColor = true;
            this.btnNetBalance.Click += new System.EventHandler(this.btnNetBalance_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSlateGray;
            this.ClientSize = new System.Drawing.Size(799, 334);
            this.Controls.Add(this.btnNetBalance);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnStatementOfAccountPDF);
            this.Controls.Add(this.btnStatementOfAccountEx);
            this.Controls.Add(this.btnOverDueInvoices);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.lblAgencyCount);
            this.Controls.Add(this.btnExportInPDF);
            this.Controls.Add(this.btnExportInExcel);
            this.Controls.Add(this.btnGetAgencyList);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CrystalReport1 CrystalReport11;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnGetAgencyList;
        private System.Windows.Forms.Button btnExportInExcel;
        private System.Windows.Forms.Button btnExportInPDF;
        private System.Windows.Forms.Label lblAgencyCount;
        
        private System.Windows.Forms.Button btnStop;
        private System.Windows.Forms.Button btnOverDueInvoices;
        private System.Windows.Forms.Button btnStatementOfAccountEx;
        private System.Windows.Forms.Button btnStatementOfAccountPDF;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btnNetBalance;
        private CachedSalesTaxInvoice cachedSalesTaxInvoice1;
        private CachedSalesTaxInvoice cachedSalesTaxInvoice2;
        private CachedSalesTaxInvoice cachedSalesTaxInvoice3;
        private CachedSalesTaxInvoice cachedSalesTaxInvoice4;
    }
}

