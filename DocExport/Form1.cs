using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;


namespace DocExport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection baglanti = new OleDbConnection();
        OleDbCommand kmt = new OleDbCommand();
        private void Form1_Load(object sender, EventArgs e)
        {
       
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Spire.DataExport.RTF.RTFExport rtfExport = new Spire.DataExport.RTF.RTFExport();
            rtfExport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            rtfExport.DataTable = this.dataGridView1.DataSource as DataTable;
            rtfExport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            Spire.DataExport.RTF.RTFStyle rtfstyle = new Spire.DataExport.RTF.RTFStyle();
            rtfstyle.FontColor = Color.Blue;
            rtfstyle.BackgroundColor = Color.LightGreen;
            rtfExport.RTFOptions.DataStyle = rtfstyle;
            rtfExport.FileName = @"..\..\06092014.doc";
            rtfExport.SaveToFile();

        }

        private void button2_Click(object sender, EventArgs e)
        {

            sorgula();
        }
        public void sorgula() 
        {

            baglanti.ConnectionString = textBox1.Text;
            kmt.CommandText = textBox2.Text;

            using (OleDbDataAdapter da = new OleDbDataAdapter())
            {
                da.SelectCommand = kmt;
                da.SelectCommand.Connection = baglanti;
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {

            Spire.DataExport.PDF.PDFExport pdfexport = new Spire.DataExport.PDF.PDFExport();
            pdfexport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            pdfexport.DataTable = this.dataGridView1.DataSource as DataTable;
            pdfexport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            pdfexport.SaveToFile("06092014.pdf");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Spire.DataExport.HTML.HTMLExport htmlexport = new Spire.DataExport.HTML.HTMLExport();
            htmlexport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            htmlexport.DataTable = this.dataGridView1.DataSource as DataTable;
            htmlexport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            htmlexport.SaveToFile("06092014.html");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Spire.DataExport.XLS.CellExport cellexport = new Spire.DataExport.XLS.CellExport();
            Spire.DataExport.XLS.WorkSheet wordshet1 = new Spire.DataExport.XLS.WorkSheet();
            wordshet1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
            wordshet1.DataTable = this.dataGridView1.DataSource as DataTable;
            wordshet1.StartDataCol = ((System.Byte)(0));
            cellexport.Sheets.Add(wordshet1);
            cellexport.ActionAfterExport = Spire.DataExport.Common.ActionType.OpenView;
            cellexport.SaveToFile("06092014.xls");


        }
    }
}
