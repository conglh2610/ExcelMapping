using Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace ExcelReaderTool
{
    public partial class frmExcelReader : Form
    {
        public frmExcelReader()
        {
            InitializeComponent();
        }
        private void btnBrowser_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtPath.Text = openFileDialog.FileName;
            }

        }

        private void excelReader (string filePath, ref IDictionary<int, string> enDict, ref IDictionary<int, string> jpDict, ref IDictionary<int, string> tpDict)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
   
            excelReader.IsFirstRowAsColumnNames = true;
            DataSet result = excelReader.AsDataSet();
            DataTable dt = result.Tables[0];
            var cols = dt.Columns.Count;
            var rows = dt.Rows.Count;
            for (int i = 0; i < rows; i++)
            {
                enDict.Add(i, dt.Rows[i][0].ToString());
                jpDict.Add(i, dt.Rows[i][1].ToString());
                tpDict.Add(i, dt.Rows[i][2].ToString());
            } 
            
        }

        private List<string> readXml()
        {
            var res = new List<string>();
            var xmlPath = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName + "\\MappingRule.xml";
            XDocument xdoc = XDocument.Load(xmlPath);
            //Run query
           foreach(var element in xdoc.Descendants("Columns"))
            {
                var column = element.Element("Value").Value;
                res.Add(column);
            }

            return res;
        }

        private void btnMapping_Click(object sender, EventArgs e)
        {
            var res = rtbSourceCode.Text;
            //var xmlStructure = readXml();
            IDictionary<int, string> enDict = new Dictionary<int, string>();
            IDictionary<int, string> jpDict = new Dictionary<int, string>();
            IDictionary<int, string> tpDict = new Dictionary<int, string>();
            excelReader(txtPath.Text, ref enDict, ref jpDict, ref tpDict);
            for (int i = 0; i < enDict.Count; i++)
            {
               res = res.Replace(jpDict[i], enDict[i]);
                
            }

            var outPutText = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName + "\\output-text.txt";
            if (!File.Exists(outPutText))
            {
                File.Create(outPutText);
            }

            File.WriteAllText(outPutText, res);
            Process.Start("notepad++.exe", outPutText);
        }
    }
}
