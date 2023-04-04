using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace SaveFileNameToExcel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            throw new NotImplementedException();
        }

        private void button1_Click2(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string[] files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*", SearchOption.AllDirectories);

                var excelFile = new FileInfo(@"C:\example.xlsx");
                using (var package = new ExcelPackage(excelFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    int row = 1;

                    foreach (string file in files)
                    {
                        worksheet.Cells[row, 1].Value = file;
                        row++;
                    }

                    package.Save();
                }
            }
        }
    }
}
