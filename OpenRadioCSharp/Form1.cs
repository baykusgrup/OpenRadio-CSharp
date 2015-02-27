using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenRadioCSharp
{
    public partial class OpenRadio : Form
    {
        public OpenRadio()
        {
            InitializeComponent();
        }

        string path = Application.StartupPath + "\\files\\urls.xlsx";

        private void Form1_Load(object sender, EventArgs e)
        {
            string yolBaglantisi = string.Concat("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=", path, ";Extended Properties='Excel 12.0 XML;HDR=YES;';");
            OleDbConnection baglanti = new OleDbConnection(yolBaglantisi);
            baglanti.Open();
            OleDbDataAdapter veriAdaptor = new OleDbDataAdapter(@"SELECT * FROM [Sheet1$]", baglanti);

            DataTable tablo = new DataTable();

            veriAdaptor.Fill(tablo);
            dataGridView1.DataSource = tablo;
            baglanti.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //Seçili satırın X koordinatı

            int xkoordinat = dataGridView1.CurrentCellAddress.X;

            //Seçili satırın Y koordinatı

            int ykoordinat = dataGridView1.CurrentCellAddress.Y;

            string str = "";

            str = dataGridView1.Rows[ykoordinat].Cells[xkoordinat].Value.ToString();
            axWindowsMediaPlayer1.URL = str;
        }
    }
}
