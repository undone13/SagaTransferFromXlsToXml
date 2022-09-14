using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelDataReader;
using System.Xml;

namespace SagaTransferFromXlsToXml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string fileName = "documentXML";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonChooseFile_Click(object sender, RoutedEventArgs e)
        {
            string filePath = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Files|*.xls;*.xlsx";
            openFileDialog1.Title = "Select a File";

            if (openFileDialog1.ShowDialog() == true)
            {
                filePath = openFileDialog1.FileName;
            }
            openFileDialog1 = null;

            if(File.Exists(fileName) == false)
            {
                CreateXML(fileName);
            }
            else
            {
                //delete file
            }

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            InvoiceClass.Antet antet = new InvoiceClass.Antet();
                            antet.ClientNume = reader.GetString(0);
                            antet.ClientCIF = reader.GetString(1);
                            antet.FacturaNumar = reader.GetString(2).Split(' ').Skip(1).FirstOrDefault(); //fact 1234 -> 1234
                            antet.FurnizorNume = reader.GetString(3);
                            antet.FurnizorCIF = reader.GetString(4);

                            CreateInvoiceAntet(fileName, antet);
                        }
                    } while (reader.NextResult());


                    //var result = reader.AsDataSet();
                    //var x = result.Tables[0];
                    ////MessageBox.Show(result.Tables[0].Rows.Count.ToString());
                }
            }
        }

        private void CreateXML(string xmlFileName) 
        {
            XmlDocument doc = new XmlDocument();

            XmlNode facturi = doc.CreateElement("Facturi");
            doc.AppendChild(facturi);

            doc.Save(xmlFileName + ".xml");
        }

        private void CreateInvoiceAntet(string xmlFileName, InvoiceClass.Antet antet)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName + ".xml");

            XmlNode facturi = doc.SelectSingleNode("Facturi");

            XmlNode factura = doc.CreateElement("Factura");
            facturi.AppendChild(factura);

            XmlNode FurnizorNume = doc.CreateElement("FurnizorNume");
            FurnizorNume.InnerText = antet.FurnizorNume;
            XmlNode FurnizorCIF = doc.CreateElement("FurnizorCIF");
            FurnizorCIF.InnerText = antet.FurnizorCIF;
            XmlNode FurnizorNrRegCom = doc.CreateElement("FurnizorNrRegCom");
            FurnizorNrRegCom.InnerText = antet.FurnizorNrRegCom;
            XmlNode FurnizorCapital = doc.CreateElement("FurnizorCapital");
            FurnizorCapital.InnerText = antet.FurnizorCapital;
            XmlNode FurnizorAdresa = doc.CreateElement("FurnizorAdresa");
            FurnizorAdresa.InnerText = antet.FurnizorAdresa;

            XmlNode FurnizorBanca = doc.CreateElement("FurnizorBanca");
            FurnizorBanca.InnerText = antet.FurnizorBanca;
            XmlNode FurnizorIBAN = doc.CreateElement("FurnizorIBAN");
            FurnizorIBAN.InnerText = antet.FurnizorIBAN;
            XmlNode FurnizorInformatiiSuplimentare = doc.CreateElement("FurnizorInformatiiSuplimentare");
            FurnizorInformatiiSuplimentare.InnerText = antet.FurnizorInformatiiSuplimentare;
            XmlNode ClientNume = doc.CreateElement("ClientNume");
            ClientNume.InnerText = antet.ClientNume;
            XmlNode ClientInformatiiSuplimentare = doc.CreateElement("ClientInformatiiSuplimentare");
            ClientInformatiiSuplimentare.InnerText = antet.ClientInformatiiSuplimentare;

            XmlNode ClientCIF = doc.CreateElement("ClientCIF");
            ClientCIF.InnerText = antet.ClientCIF;
            XmlNode ClientNrRegCom = doc.CreateElement("ClientNrRegCom");
            ClientNrRegCom.InnerText = antet.ClientNrRegCom;
            XmlNode ClientJudet = doc.CreateElement("ClientJudet");
            ClientJudet.InnerText = antet.ClientJudet;
            XmlNode ClientAdresa = doc.CreateElement("ClientAdresa");
            ClientAdresa.InnerText = antet.ClientAdresa;
            XmlNode ClientBanca = doc.CreateElement("ClientBanca");
            ClientBanca.InnerText = antet.ClientBanca;

            XmlNode ClientIBAN = doc.CreateElement("ClientIBAN");
            ClientIBAN.InnerText = antet.ClientIBAN;
            XmlNode FacturaNumar = doc.CreateElement("FacturaNumar");
            FacturaNumar.InnerText = antet.FacturaNumar;
            XmlNode FacturaData = doc.CreateElement("FacturaData");
            FacturaData.InnerText = antet.FacturaData;
            XmlNode FacturaScadenta = doc.CreateElement("FacturaScadenta");
            FacturaScadenta.InnerText = antet.FacturaScadenta;
            XmlNode FacturaTaxareInversa = doc.CreateElement("FacturaTaxareInversa");
            FacturaTaxareInversa.InnerText = antet.FacturaTaxareInversa;

            XmlNode FacturaTVAIncasare = doc.CreateElement("FacturaTVAIncasare");
            FacturaTVAIncasare.InnerText = antet.FacturaTVAIncasare;
            XmlNode FacturaInformatiiSuplimentare = doc.CreateElement("FacturaInformatiiSuplimentare");
            FacturaInformatiiSuplimentare.InnerText = antet.FacturaInformatiiSuplimentare;
            XmlNode FacturaMoneda = doc.CreateElement("FacturaMoneda");
            FacturaMoneda.InnerText = antet.FacturaMoneda;


            factura.AppendChild(FurnizorNume);
            factura.AppendChild(FurnizorCIF);
            factura.AppendChild(FurnizorNrRegCom);
            factura.AppendChild(FurnizorCapital);
            factura.AppendChild(FurnizorAdresa);

            factura.AppendChild(FurnizorBanca);
            factura.AppendChild(FurnizorIBAN);
            factura.AppendChild(FurnizorInformatiiSuplimentare);
            factura.AppendChild(ClientNume);
            factura.AppendChild(ClientInformatiiSuplimentare);

            factura.AppendChild(ClientCIF);
            factura.AppendChild(ClientNrRegCom);
            factura.AppendChild(ClientJudet);
            factura.AppendChild(ClientAdresa);
            factura.AppendChild(ClientBanca);

            factura.AppendChild(ClientIBAN);
            factura.AppendChild(FacturaNumar);
            factura.AppendChild(FacturaData);
            factura.AppendChild(FacturaScadenta);
            factura.AppendChild(FacturaTaxareInversa);

            factura.AppendChild(FacturaTVAIncasare);
            factura.AppendChild(FacturaInformatiiSuplimentare);
            factura.AppendChild(FacturaMoneda);

            doc.Save(xmlFileName + ".xml");  
        }
    }
}
