using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
namespace OLEDB_Excel_App
{
    public partial class Form1 : Form
    {
        string filePath;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            filePath = DosyaAcVePathBul();

            OleDbConnection oleDbConnection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={filePath}; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            oleDbConnection.Open();

            DataTable sheets = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            comboBoxSheets.Items.Clear();
            foreach (DataRow sheet in sheets.Rows)
            {
                string sheetName = sheet["TABLE_NAME"].ToString();
                comboBoxSheets.Items.Add(sheetName);
            }
            if (comboBoxSheets.Items.Count > 0)
                comboBoxSheets.SelectedIndex = 0;
            oleDbConnection.Close();

        }
        public string DosyaAcVePathBul()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
            openFileDialog.ShowDialog();
            groupBox1.Text = openFileDialog.FileName;

            return groupBox1.Text.ToString();
        }

        private void comboBoxSheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedSheet = comboBoxSheets.SelectedItem.ToString();

            OleDbConnection oleDbConnection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={groupBox1.Text}; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            oleDbConnection.Open();

            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter($"SELECT * FROM [{selectedSheet}]", oleDbConnection);

            DataTable dataTable = new DataTable();

            oleDbDataAdapter.Fill(dataTable);

            dataGridView1.DataSource = dataTable.DefaultView;

            oleDbConnection.Close();
        }

        private void btnApplyFilter_Click(object sender, EventArgs e)
        {
            string filterQuery = textBoxFilter.Text;

            if (string.IsNullOrEmpty(filterQuery))
            {
                MessageBox.Show("Lütfen bir filtre sorgusu girin.");
                return;
            }
            
            string selectedSheet = comboBoxSheets.SelectedItem.ToString();

            OleDbConnection oleDbConnection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={groupBox1.Text}; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            oleDbConnection.Open();

            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter($"SELECT * FROM [{selectedSheet}] WHERE {filterQuery}", oleDbConnection);
            
            DataTable dataTable = new DataTable();

            oleDbDataAdapter.Fill(dataTable);

            dataGridView1.DataSource = dataTable.DefaultView;

            oleDbConnection.Close();

        }
    }
}


