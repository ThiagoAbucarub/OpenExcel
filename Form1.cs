using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenExcel
{
    public partial class Form1 : Form
    {
        Excel.Application excelApp;
        Workbook excelWorkBook;
        Worksheet excelWorkSheet;
        System.Data.DataTable dataTable;

        string caminhoArquivo, extensaoArquivo, stringConexao, salvarXML, id1;

        int i = 0;
        int totalLinhas = 1;

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnOpenFiles_Click(object sender, EventArgs e)
        {
            AbrirArquivo();
            LerArquivo();
        }

       private void BtnSaveXml_Click(object sender, EventArgs e)
        {
            GerarXml();
            LimparMemoria();
        }

        private void AbrirArquivo()
        {
            OpenFileDialog abrirArquivo = new OpenFileDialog();
            abrirArquivo.Title = "Abrir Arquivo";
            abrirArquivo.Filter = "Arquivo Excel |*.xls;*.xlsx";

            if (abrirArquivo.ShowDialog() == DialogResult.OK)
            {
                caminhoArquivo = abrirArquivo.FileName;
                extensaoArquivo = Path.GetExtension(caminhoArquivo);
                excelApp = new Excel.Application();
                excelWorkBook = excelApp.Workbooks.Open(caminhoArquivo);
                txtNameFile.Text = caminhoArquivo;
                excelWorkSheet = (Worksheet)excelWorkBook.Sheets[1];
            }
        }

        private void LerArquivo()
        {
            if (extensaoArquivo == ".xls" || extensaoArquivo == ".XLS")
                stringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + caminhoArquivo + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            else if (extensaoArquivo == ".xlsx" || extensaoArquivo == ".XLSX")
                stringConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminhoArquivo + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";

            OleDbConnection connection = new OleDbConnection(stringConexao);
            OleDbCommand commmand = new OleDbCommand();
            commmand.Connection = connection;
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(commmand);
            dataTable = new System.Data.DataTable();
            connection.Open();

            System.Data.DataTable dataSheet = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string sheetName = dataSheet.Rows[0]["table_name"].ToString();
            commmand.CommandText = "select * from [" + sheetName + "]";
            dataAdapter.SelectCommand = commmand;
            dataAdapter.Fill(dataTable);
            connection.Close();

            dataGridView1.DataSource = dataTable;

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);
            totalLinhas = dataTable.Rows.Count;
        }

        private void GerarXml()
        {
            XmlDocument xml = new XmlDocument();
            XmlDeclaration xmlDeclaration = xml.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlElement root = xml.DocumentElement;
            xml.InsertBefore(xmlDeclaration, root);

            XmlElement dados = xml.CreateElement("Dados");
            xml.AppendChild(dados);

            for (i = 0; i < totalLinhas; i++)
            {
                id1 = dataTable.Rows[i].ItemArray[1].ToString();

                XmlElement info = xml.CreateElement("Informações");

                XmlElement id = xml.CreateElement("Id");
                id.InnerText = dataTable.Rows[i].ItemArray[0].ToString();
                info.AppendChild(id);

                XmlElement nome = xml.CreateElement("Nome");
                nome.InnerText = dataTable.Rows[i].ItemArray[1].ToString();
                info.AppendChild(nome);

                XmlElement telefone = xml.CreateElement("Telefone");
                telefone.InnerText = dataTable.Rows[i].ItemArray[2].ToString();
                info.AppendChild(telefone);

                XmlElement sexo = xml.CreateElement("Sexo");
                sexo.InnerText = dataTable.Rows[i].ItemArray[3].ToString();
                info.AppendChild(sexo);

                dados.AppendChild(info);
                xml.DocumentElement.AppendChild(info);
            }

            SalvarArquivo();

            xml.Save(salvarXML);

        }

        private void SalvarArquivo()
        {
            SaveFileDialog salvarArquivo = new SaveFileDialog();
            salvarArquivo.Title = "Save As";
            salvarArquivo.Filter = "XML Files |*.xml";
            salvarArquivo.FilterIndex = 0;
            salvarArquivo.FileName = "Sample_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            salvarArquivo.InitialDirectory = @"c:\xml";
            salvarArquivo.RestoreDirectory = true;

            if (salvarArquivo.ShowDialog() == DialogResult.OK)
            {
                salvarXML = salvarArquivo.FileName;
                MessageBox.Show("Arquivo salvo!");
            }
        }

        private void LimparMemoria()
        {
            excelWorkBook.Close(true, null, null);
            excelApp.Quit();

            Marshal.ReleaseComObject(excelWorkSheet);
            Marshal.ReleaseComObject(excelWorkBook);
            Marshal.ReleaseComObject(excelApp);
        }
        
    }
}
