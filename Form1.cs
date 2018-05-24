using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenExcel
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Workbook xlWorkBook;
        Worksheet xlWorkSheet;
        Range range;

        public Form1()
        {
            InitializeComponent();
        }

        public void openFile()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Abrir Arquivo";
            openFile.Filter = "Arquivo Excel |*.xls;*.xlsx";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string caminhoArquivo = openFile.FileName;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(caminhoArquivo);
                txtNameFile.Text = caminhoArquivo;
            }

        }

        private void btnOpenFiles_Click(object sender, EventArgs e)
        {
            openFile();

            xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];

            string str = "";
            int rCnt = 2;
            int cCnt = 1;
            int rw = 0;
            int cl = 0;

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            List<AcctOpngInstr> _acctOpngInstr = new List<AcctOpngInstr>();

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    AcctOpngInstr acctOpngInstr = new AcctOpngInstr();
                    acctOpngInstr.Nome = (range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString();
                    _acctOpngInstr.Add(acctOpngInstr);

                }
            }
            
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            GerarXml();
        }
    
        private void GerarXml()
        {
            XmlDocument xml = new XmlDocument();
            XmlDeclaration xmlDeclaration = xml.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xml.DocumentElement;
            xml.InsertBefore(xmlDeclaration, root);

            XmlElement Dados = xml.CreateElement("Dados");
            xml.AppendChild(Dados);

            XmlElement Nome = xml.CreateElement("Nome");
            Nome.AppendChild(xml.CreateTextNode(lblName2.Text));
            Dados.AppendChild(Nome);

            XmlElement Phone = xml.CreateElement("Telefone");
            Phone.AppendChild(xml.CreateTextNode(lblPhone2.Text));
            Dados.AppendChild(Phone);

            XmlElement Email = xml.CreateElement("Sexo");
            Email.AppendChild(xml.CreateTextNode(lblEmail2.Text));
            Dados.AppendChild(Email);


            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Title = "Save As";
            saveFile.Filter = "XML Files |*.xml";
            saveFile.FilterIndex = 0;
            saveFile.FileName = "Sample_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            saveFile.InitialDirectory = @"c:\xml";
            saveFile.RestoreDirectory = true;

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                xml.Save(saveFile.FileName);
                MessageBox.Show("Arquivo salvo!");
            }
            
        }
    }
}
