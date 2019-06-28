using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WindowsFormsApp4
{
    public partial class ExcelForm : Form
    {
        private static int _counter = 0;
        public ExcelForm()
        {
            InitializeComponent();
        }

        private void newToolStripButton_Click(object sender, EventArgs e)
        {
            CreateForm();
        }
        public static ExcelForm CreateForm()
        {
            var form = new ExcelForm();
            form.Text = "New Excel sheet " + ++_counter;
            ExcelApplication.Instance.ApplicationContext.MainForm = form;
            form.Show();

            return form;
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                BindDataCSV(fileName);
            }
        }

        
        private void BindDataCSV(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath);
            string[] headerLables = lines[0].Split(',');

            for(int x = 0; x < lines.Length; x++ )
            {
                int iCounter = 0;
                string[] dataWords = lines[x].Split(',');
                dataGridView1.Rows.Add();
                foreach (string headerWord in headerLables)
                {


                    dataGridView1.Rows[x].Cells[iCounter].Value = dataWords[iCounter];

                   //FutureWork area.  Purpose of future work:
                   //   if the row cells have data in them, start importing data into the next row whose first column has 'null' value.

                    //int rowsize = dataGridView1.Rows.Count;
                    //int a = 0;
                    //for (; a < rowsize; a++)
                    //{
                    //    if (dataGridView1.Rows[a].Cells[x].Value == null)
                    //    {
                                                        
                    //    }
                    //}

                    iCounter++;
                }
            }
        }

       

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Convert.ToString(dataGridView1.Rows[0].Cells[0].Value));
        }

        private void saveToolStripButton_Click_1(object sender, EventArgs e)
        {
           
            SaveToXLS();

        }
        private void SaveToXLS()
        {
            //save dialog formatting
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "";
            saveFileDialog1.Filter = "Excel File | *.xls";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //creating excel file 
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);

                //reading colums/rows
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                //saving file and exit
                ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.Quit();
            }
            MessageBox.Show($"Your file was successfully saved as : {saveFileDialog1.FileName}");
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateForm();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var fileName = openFileDialog1.FileName;
                BindDataCSV(fileName);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveToXLS();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            MessageBox.Show("Hello?\n" +
                "This is Jimmy. " +
                "I am a junior developer at 00 Korea Inc. " +
                "Thank you for your time.", this.Text = "About");


        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Please, seek help from google or stackoverflow.\n" +
                "Thank you for understanding :)", this.Text = "Help"); 
        }
    }
}
