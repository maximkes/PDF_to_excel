using System;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using IronXL;


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

        private void SaveToExcel(List<string> Text, string path, string name)
        {
            name = GenerateName(path, name);
            var newXLFile = WorkBook.Create(ExcelFileFormat.XLSX);
            newXLFile.Metadata.Title = "IronXL New File";
            var newWorkSheet = newXLFile.CreateWorkSheet("1stWorkSheet");
            //newWorkSheet["A1"].Value = "Hello World";
            int N=Text.Count;
            for(int i =0; i< Text.Count; i++)
            {
                newWorkSheet[("A"+(i+1).ToString())].Value = Text[i];
                textBox1.Text = ((i+1).ToString() + "/" + N.ToString() + ": " + Text[i]);
            }
            textBox1.Text = "���� " + name + ".xlsx" + " ������� ������";
            string new_path = path + "\\" + name + ".xlsx";
            newXLFile.SaveAs(new_path);

            
        }

        private string GetFileName(string path)
        {
            string res = "";
            res = path.Split('\\')[path.Split('\\').Length-1];
            res = res.Replace(".pdf", "");
            return res;
        }
        
        private string GetFilePath(string path)
        {
            return path.Substring(0, path.LastIndexOf('\\'));
        }

        private void button1_Click(object sender, EventArgs e)
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
            button1.Enabled = false;
            var TextList = GetText(filePath);
            for(int i=0; i<TextList.Count; i++)
            {
                string L = "";
                foreach(var x in TextList[i].Split("\n"))
                    L = L + x;
                TextList[i] = L;
                int a = (TextList[i]).Length;
            }
            SaveToExcel(TextList, GetFilePath(filePath), GetFileName(filePath));
            button1.Enabled = true;
        }

        private string GenerateName(string path, string name)
        {
            string res="";
            var directory = new DirectoryInfo(path);
            FileInfo[] files = directory.GetFiles();
            List<string> FileNames = new List<string>();
            foreach (FileInfo file in files)
            {
                FileNames.Add(file.Name);
            }
            if (!FileNames.Contains(name + ".xlsx")) 
                return name;
            int n = 1;
            while (FileNames.Contains(name + "(" + n.ToString() + ").xlsx")) 
                n++;
            //string s = name + "(" + n.ToString() + ").xlsx";
            return (name + "(" + n.ToString() + ")");
        }
    }
}