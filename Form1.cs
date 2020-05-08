using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using CrystalDecisions.CrystalReports.Engine;

namespace CrystalReportGrouping
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /*SqlConnection sqlcon = new SqlConnection(@"Data Source=4514101956-MY\SQLEXPRESS;Initial Catalog=DB1;Integrated Security=True");
            sqlcon.Open();
            SqlDataAdapter sqlDA = new SqlDataAdapter("SalesView", sqlcon);
            sqlDA.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataSet dsProd = new DataSet();
            sqlDA.Fill(dsProd);
            sqlcon.Close();
            crptGroupReport crptProduct = new crptGroupReport();

            crptProduct.Database.Tables["Sales"].SetDataSource(dsProd.Tables[0]);
           
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = crptProduct;*/

            crptGroupReport crpt2 = new crptGroupReport();
           // crpt2.SetParameterValue("@Price", "65");
            crystalReportViewer1.ReportSource = null;
            crystalReportViewer1.ReportSource = crpt2;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ReportDocument cryRpt = new ReportDocument();

            cryRpt.Load(@"C:\Users\Administrator\Documents\visual studio 2010\Projects\ProductCrystalRept\CrystalReportGrouping\crptGroupReport.rpt");

            crystalReportViewer1.ReportSource = cryRpt;

            crystalReportViewer1.Refresh();

            cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelWorkbook, @"C:\ASD.xls");

            MessageBox.Show("Exported Successful");
        }
    }
}
