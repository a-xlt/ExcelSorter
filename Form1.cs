using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelSorter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog open = new OpenFileDialog();
           
            
            if (open.ShowDialog() == DialogResult.OK)
            {
                  
                string inputFilePath = open.FileName;

                string outputName = "output.xlsx";

                if (!string.IsNullOrWhiteSpace(textBox1.Text.Trim()))
                {
                    outputName = textBox1.Text.Trim() + ".xlsx";
                }

                string outputFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), outputName);

                
                Dictionary<string, int> valueCounts = new Dictionary<string, int>();

                 using (var package = new ExcelPackage(new FileInfo(inputFilePath)))
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                     for (int row = 1; row <= rowCount; row++)
                    {
                        string value = worksheet.Cells[row, 1].Text;
                        if (valueCounts.ContainsKey(value))
                        {
                            valueCounts[value]++;
                        }
                        else
                        {
                            valueCounts[value] = 1;
                        }
                    }
                }

                 using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("الباركودات");

                     worksheet.Cells[1, 1].Value = "الباركود";
                    worksheet.Cells[1, 2].Value = "العدد";

                    int row = 2;
                     foreach (var kvp in valueCounts)
                    {
                        worksheet.Cells[row, 1].Value = kvp.Key;
                        worksheet.Cells[row, 2].Value = kvp.Value;
                        row++;
                    }

                     package.SaveAs(new FileInfo(outputFilePath));
                }



                MessageBox.Show("تمت العملية بنجاح");
                textBox1.Text=string.Empty;

            }



        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
