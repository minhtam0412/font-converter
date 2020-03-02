using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Security.Permissions;
using System.Windows.Forms;

namespace ExcelConverter
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class Form1 : Form
    {
        //https://vietcatholicsydney.net/Convert/index.htm
        public Form1()
        {
            InitializeComponent();
        }

        private DataTable dtbDestination;

        public DataTable ReadExcelFile(string filePath)
        {
            DataTable dtexcel = new DataTable();
            bool hasHeaders = false;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=0\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //Looping a first Sheet of Xl File
            DataRow schemaRow = schemaTable.Rows[0];
            string sheet = schemaRow["TABLE_NAME"].ToString();
            if (!sheet.EndsWith("_"))
            {
                string query = "SELECT  * FROM [" + sheet + "]";
                OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                dtexcel.Locale = CultureInfo.CurrentCulture;
                daexcel.Fill(dtexcel);
            }

            conn.Close();
            return dtexcel;
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSourcePath.Text))
            {
                MessageBox.Show("Vui lòng chọn tập tin cần chuyển đổi!");
                return;
            }
            var dtbSource = this.ReadExcelFile(txtSourcePath.Text.Trim());
            dataGridView1.DataSource = dtbSource;
            dtbDestination = dtbSource.Copy();
            foreach (DataRow row in dtbDestination.Rows)
            {
                foreach (DataColumn column in dtbDestination.Columns)
                {
                    var src = Convert.ToString(row[column]);
                    if (!string.IsNullOrEmpty(src))
                    {
                        {
                            var rsl = webBrowser1.Document.InvokeScript("Convert", new object[] { src, 3, 5 });
                            Debug.WriteLine(src + " -> " + rsl);

                            row[column] = chkReplaceNewLine.Checked ? rsl.ToString().Replace("\n", " ").Replace("\r\n", " ") : rsl;
                        }
                    }
                }
            }

            dataGridView2.DataSource = dtbDestination;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            webBrowser1.ObjectForScripting = this;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string currentPath = Directory.GetCurrentDirectory();
            webBrowser1.Navigate(Path.Combine(currentPath, "HTMLPage1.html"));
            webBrowser1.ScriptErrorsSuppressed = true;
        }

        private void btnExportResult_Click(object sender, EventArgs e)
        {
            if (dtbDestination == null || dtbDestination.Rows.Count == 0)
            {
                MessageBox.Show("Không có kết quả để xuất!");
                return;
            }

            saveFileDialog1.Title = "Lưu kết quả";
            saveFileDialog1.Filter = "Excel 2007 above|*.xlsx|Excel 2003|*.xls";
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.CheckPathExists = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFileDialog1.FileName;
                XLWorkbook xlWorkbook = new XLWorkbook();
                xlWorkbook.Worksheets.Add(dtbDestination, "Sheet1");
                xlWorkbook.SaveAs(fileName);
            }
        }

        private void btnChooseFile_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel 2007 above|*.xlsx|Excel 2003|*.xls";
            openFileDialog1.Multiselect = false;
            openFileDialog1.Title = "Chọn tập tin";
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtSourcePath.Text = openFileDialog1.FileName;
            }
        }
    }
}