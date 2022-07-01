using System;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
//using System.IO;

namespace ChooseFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "PDF Files|*.pdf";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            var TextList = GetText(filePath);
            for(int i=0; i<TextList.Count; i++)
            {
                TextList[i] = TextList[i].Replace("\n", "");
                listBox1.Items.Insert(i,TextList[i]);

            }
            textBox1.Text = TextList[0];
            
            //MessageBox.Show(TextList.ToString(), "File Content at path: " + filePath, MessageBoxButtons.OK);

            //SaveToExcel(TextList, filePath);
        }

        private List<string> GetText(string path)
        {
            var res = new List<string>();
            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    res.Add(PdfTextExtractor.GetTextFromPage(reader, i));
                }

                return res;
            }
        }

        private void SaveToExcel(List<string> Text, string path)
        {
            Excel.Application excelApp = new Excel.Application();

            // —делать приложение Excel видимым
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;
            // ”становить заголовки столбцов в €чейках
            for(int i=0;i<Text.Count;i++)
            {
                workSheet.Cells[i, "A"] = Text[i];
            }
            
            excelApp.DisplayAlerts = false;
            workSheet.SaveAs(path.Replace(".pdf", ".xlsx"));

            excelApp.Quit();

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}