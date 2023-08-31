using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

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

            System.Data.DataTable sheets = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            comboBoxSheets.Items.Clear();
            foreach (DataRow sheet in sheets.Rows)
            {
                string sheetName = sheet["TABLE_NAME"].ToString();
                comboBoxSheets.Items.Add(sheetName);
            }
            if (comboBoxSheets.Items.Count > 0)
                comboBoxSheets.SelectedIndex = 0;
            /*
            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("SELECT * FROM [Data$]",oleDbConnection);

            DataTable dataTable = new DataTable();

            oleDbDataAdapter.Fill(dataTable);

            dataGridView1.DataSource = dataTable.DefaultView;
            */
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

            System.Data.DataTable dataTable = new System.Data.DataTable();

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
            
            //OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter($"{filterQuery}", oleDbConnection);

            System.Data.DataTable dataTable = new System.Data.DataTable();

            oleDbDataAdapter.Fill(dataTable);

            dataGridView1.DataSource = dataTable.DefaultView;

            oleDbConnection.Close();

        }
        //tekrar incelenmek üzere silinmedi.
       /*
        * private void btnSaveAs_Click(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource == null)
            {
                MessageBox.Show("Önce bir sorgu sonucu alın.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string excelFilePath = groupBox1.Text; // Excel dosyasının yolu ve adı
            string worksheetName = "sql"; // Sayfa adı

            try
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                Excel._Worksheet worksheet = null;

                foreach (Excel._Worksheet sheet in workbook.Sheets)
                {
                    if (sheet.Name == worksheetName)
                    {
                        worksheet = sheet;
                        break;
                    }
                }

                if (worksheet == null)
                {
                    MessageBox.Show("Sayfa bulunamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workbook.Close();
                    excelApp.Quit();
                    return;
                }

                DataView queryResultView = (DataView)dataGridView1.DataSource;
                System.Data.DataTable queryResultTable = queryResultView.ToTable();

                int row = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;

                for (int i = 0; i < queryResultTable.Rows.Count; i++)
                {
                    for (int j = 0; j < queryResultTable.Columns.Count; j++)
                    {
                        worksheet.Cells[row + i, j + 1] = queryResultTable.Rows[i][j].ToString();
                    }
                }

                workbook.Save();
                workbook.Close();
                excelApp.Quit();

                MessageBox.Show("İşlem Başarılı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /*
        if (dataGridView1.DataSource == null)
        {
            MessageBox.Show("Önce bir sorgu sonucu alın.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Excel Dosyası |*.xlsx";
        saveFileDialog.Title = "Sorgu Sonucunu Excel'e Kaydet";
        saveFileDialog.ShowDialog();

        if (saveFileDialog.FileName != "")
        {
            try
            {
                DataView queryResultView = (DataView)dataGridView1.DataSource;
                DataTable queryResultTable = queryResultView.ToTable();

                using (OleDbConnection oleDbConnection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={saveFileDialog.FileName}; Extended Properties='Excel 12.0 xml;HDR=YES;'"))
                {
                    oleDbConnection.Open();

                    string tableName = "Sorgu Sonucu";
                    string createTableQuery = $"CREATE TABLE [{tableName}] (";

                    createTableQuery += "[EEID] INTEGER,";
                    createTableQuery += "[Full Name] TEXT,";
                    createTableQuery += "[Job Title] TEXT,";
                    createTableQuery += "[Department] TEXT,";
                    createTableQuery += "[Business Unit] TEXT,";
                    createTableQuery += "[Gender] TEXT,";
                    createTableQuery += "[Ethnicity] TEXT,";
                    createTableQuery += "[Age] INTEGER,";
                    createTableQuery += "[Hire Date] DATETIME,";
                    createTableQuery += "[Annual Salary] DOUBLE,";
                    createTableQuery += "[Bonus %] DOUBLE,";
                    createTableQuery += "[Country] TEXT,";
                    createTableQuery += "[City] TEXT,";
                    createTableQuery += "[Exit Date] DATETIME)";

                    //createTableQuery = createTableQuery.TrimEnd(',', ' ') + ")";
                    using (OleDbCommand createTableCommand = new OleDbCommand(createTableQuery, oleDbConnection))
                    {
                        createTableCommand.ExecuteNonQuery();
                    }

                    foreach (DataRow row in queryResultTable.Rows)
                    {
                        string insertQuery = $"INSERT INTO [{tableName}] ([EEID], [Full Name], [Job Title], [Department], [Business Unit], [Gender], [Ethnicity], [Age], [Hire Date], [Annual Salary], [Bonus %], [Country], [City], [Exit Date]) VALUES (";

                        for (int i = 0; i < row.ItemArray.Length; i++)
                        {
                            object item = row.ItemArray[i];

                            if (queryResultTable.Columns[i].DataType == typeof(int))
                            {
                                insertQuery += $"{item}, ";
                            }
                            else if (queryResultTable.Columns[i].DataType == typeof(double))
                            {
                                insertQuery += $"{item.ToString().Replace(',', '.')}, ";
                            }
                            else if (queryResultTable.Columns[i].DataType == typeof(DateTime))
                            {
                                DateTime dateValue = (DateTime)item;
                                insertQuery += $"'{dateValue.ToString("yyyy-MM-dd HH:mm:ss")}', ";
                            }
                            else
                            {
                                insertQuery += $"'{item.ToString().Replace("'", "''")}', ";
                            }
                        }

                        insertQuery = insertQuery.TrimEnd(',', ' ') + ")";

                        using (OleDbCommand insertCommand = new OleDbCommand(insertQuery, oleDbConnection))
                        {
                            insertCommand.ExecuteNonQuery();
                        }
                    }

                    oleDbConnection.Close();

                    MessageBox.Show("İşlem Başarılı.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        */

    }

}


