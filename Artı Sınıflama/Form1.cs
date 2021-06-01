using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Artı_Sınıflama
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        DataTable dt = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btn_al_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile1 = new OpenFileDialog
            {
                Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                Title = "Veri Excel'ini seçiniz..."
            };
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openfile1.FileName;
            }

            Excel.Application oXL = new Excel.Application(); //hmm demek nuget paketten bulmak gerekiyormuş seni ve sonrada öyle using Excel diyerek kullanmak gerekiyormuş
            if (textBox1.Text == string.Empty)
            {
                return;
            }
            else
            {
                Excel.Workbook oWB = oXL.Workbooks.Open(textBox1.Text); // hata burada oluşuyor demek

                List<string> liste = new List<string>();
                foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                {
                    liste.Add(oSheet.Name);
                }
                oWB.Close();
                oXL.Quit();
                oWB = null;
                oXL = null;
                metroGrid2.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                textBox2.Text = metroGrid2.Rows[0].Cells[0].Value.ToString();

                OleDbCommand komut = new OleDbCommand();
                string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                OleDbConnection conn = new OleDbConnection(pathconn);
                OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                MyDataAdapter.Fill(dt);
                metroGrid1.DataSource = dt;

                metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
                metroLabel3.Text = metroGrid1.Rows.Count.ToString();
            }
        }

        private void btn_ver_Click(object sender, EventArgs e)
        {
            Excel.Application uyg = new Excel.Application();
            uyg.Visible = true;
            Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet sheet1 = (Excel.Worksheet)kitap.Sheets[1];
            for (int i = 0; i < metroGrid1.Columns.Count; i++)
            {
                Excel.Range myRange = (Excel.Range)sheet1.Cells[1, i + 1];
                myRange.Value2 = metroGrid1.Columns[i].HeaderText;
            }

            for (int i = 0; i < metroGrid1.Columns.Count; i++)
            {
                for (int j = 0; j < metroGrid1.Rows.Count; j++)
                {
                    Excel.Range myRange = (Excel.Range)sheet1.Cells[j + 2, i + 1];
                    myRange.Value2 = metroGrid1[i, j].Value;
                }
            }
        }

        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            DataView dv = dt.DefaultView;
            dv.RowFilter = "Ücretler LIKE '" + metroTextBox1.Text + "%'";
            metroGrid1.DataSource = dv;
            metroLabel3.Text = metroGrid1.Rows.Count.ToString();
        }
    }
}
