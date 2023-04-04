using OfficeOpenXml;
using System;
using System.IO;
using System.Windows.Forms;

namespace SaveFileNameToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "选择文件夹";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click1);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 41);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(776, 397);
            this.textBox1.TabIndex = 1;
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "SaveFileNameToExcel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void button1_Click1(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                // 检查目录是否存在
                if (Directory.Exists(folderBrowserDialog1.SelectedPath))
                {
                    // 检查目录是否可访问
                    var directorySecurity = Directory.GetAccessControl(folderBrowserDialog1.SelectedPath);
                    if (!directorySecurity.AreAccessRulesProtected)
                    {
                        string[] files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*", SearchOption.AllDirectories);

                        // 使用 SaveFileDialog 类让用户选择文件保存的位置和名称
                        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                        saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                        saveFileDialog1.FilterIndex = 1;
                        saveFileDialog1.RestoreDirectory = true;

                        if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            var excelFile = new FileInfo(saveFileDialog1.FileName);
                            using (var package = new ExcelPackage(excelFile))
                            {
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");
                                int row = 1;

                                foreach (string file in files)
                                {
                                    // 检查文件是否存在
                                    if (File.Exists(file))
                                    {
                                        // 检查文件是否可读
                                        var attributes = File.GetAttributes(file);
                                        if ((attributes & FileAttributes.ReadOnly) != FileAttributes.ReadOnly)
                                        {
                                            worksheet.Cells[row, 1].Value = file;
                                            textBox1.AppendText(file + Environment.NewLine);
                                            row++;
                                        }
                                    }
                                }
                                package.Save();
                                DialogResult result = MessageBox.Show("Excel文件保存成功。您想要打开文件吗？", "成功", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                                if (result == DialogResult.Yes)
                                {
                                    System.Diagnostics.Process.Start(excelFile.FullName);
                                }

                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("目录受保护，无法访问。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("目录不存在。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }



        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
    }
}
