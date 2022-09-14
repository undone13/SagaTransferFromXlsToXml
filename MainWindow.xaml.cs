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
            else
            {
                MessageBox.Show("Nu s-a putut importa fisierul");
                return;
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


            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        reader.Read();
                        do
                        {
                            while (reader.Read())
                            {
                                InvoiceClass.Antet antet = new InvoiceClass.Antet();
                                antet.ClientNume = reader.GetString(0);
                                antet.ClientCIF = reader.GetString(1);
                                antet.FacturaNumar = reader.GetString(2);  //reader.GetString(2).Split(' ').Skip(1).FirstOrDefault(); //fact 1234 -> 1234
                                antet.FurnizorNume = reader.GetString(3);
                                antet.FurnizorCIF = reader.GetString(4);
                                antet.FacturaData = reader.GetString(5);
                                antet.FacturaMoneda = reader.GetString(9);

                                InvoiceClass.Linie line = new InvoiceClass.Linie();
                                line.LinieNrCrt = "1";
                                line.Descriere = "Marfa";
                                line.Cantitate = "1";
                                line.Pret = reader.GetDouble(8).ToString();
                                line.Valoare = line.Pret;
                                double TVAproc = reader.GetDouble(7) / reader.GetDouble(6) * 100f;
                                line.ProcTVA = Convert.ToInt32(TVAproc).ToString();
                                line.TVA = reader.GetDouble(7).ToString();
                                line.Cont = "371";

                                CreateInvoice(fileName, antet, line);
                            }
                        } while (reader.NextResult());

                    }
                    MessageBox.Show("Fisierul s-a creat cu succes!");
                    
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void CreateXML(string xmlFileName) 
        {
            XmlDocument doc = new XmlDocument();

            XmlNode facturi = doc.CreateElement("Facturi");
            doc.AppendChild(facturi);

            doc.Save(xmlFileName + ".xml");
        }

        private void CreateInvoice(string xmlFileName, InvoiceClass.Antet antet, InvoiceClass.Linie line)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlFileName + ".xml");

            XmlNode facturi = doc.SelectSingleNode("Facturi");

            XmlNode factura = doc.CreateElement("Factura");
            facturi.AppendChild(factura);

            #region Antet
            XmlNode antetNode = doc.CreateElement("Antet");
            factura.AppendChild(antetNode);

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


            antetNode.AppendChild(FurnizorNume);
            antetNode.AppendChild(FurnizorCIF);
            antetNode.AppendChild(FurnizorNrRegCom);
            antetNode.AppendChild(FurnizorCapital);
            antetNode.AppendChild(FurnizorAdresa);

            antetNode.AppendChild(FurnizorBanca);
            antetNode.AppendChild(FurnizorIBAN);
            antetNode.AppendChild(FurnizorInformatiiSuplimentare);
            antetNode.AppendChild(ClientNume);
            antetNode.AppendChild(ClientInformatiiSuplimentare);

            antetNode.AppendChild(ClientCIF);
            antetNode.AppendChild(ClientNrRegCom);
            antetNode.AppendChild(ClientJudet);
            antetNode.AppendChild(ClientAdresa);
            antetNode.AppendChild(ClientBanca);

            antetNode.AppendChild(ClientIBAN);
            antetNode.AppendChild(FacturaNumar);
            antetNode.AppendChild(FacturaData);
            antetNode.AppendChild(FacturaScadenta);
            antetNode.AppendChild(FacturaTaxareInversa);

            antetNode.AppendChild(FacturaTVAIncasare);
            antetNode.AppendChild(FacturaInformatiiSuplimentare);
            antetNode.AppendChild(FacturaMoneda);
            #endregion

            #region Linie
            XmlNode detalii = doc.CreateElement("Detalii");
            factura.AppendChild(detalii);
            XmlNode continut = doc.CreateElement("Continut");
            detalii.AppendChild(continut);
            XmlNode linie = doc.CreateElement("Linie");
            continut.AppendChild(linie);

            XmlNode LinieNrCrt = doc.CreateElement("LinieNrCrt");
            LinieNrCrt.InnerText = line.LinieNrCrt;
            XmlNode Descriere = doc.CreateElement("Descriere");
            Descriere.InnerText = line.Descriere;
            XmlNode CodArticolFurnizor = doc.CreateElement("CodArticolFurnizor");
            CodArticolFurnizor.InnerText = line.CodArticolFurnizor;
            XmlNode CodArticolClient = doc.CreateElement("CodArticolClient");
            CodArticolClient.InnerText = line.CodArticolClient;
            XmlNode CodBare = doc.CreateElement("CodBare");
            CodBare.InnerText = line.CodBare;

            XmlNode InformatiiSuplimentare = doc.CreateElement("InformatiiSuplimentare");
            InformatiiSuplimentare.InnerText = line.InformatiiSuplimentare;
            XmlNode UM = doc.CreateElement("UM");
            UM.InnerText = line.UM;
            XmlNode Cantitate = doc.CreateElement("Cantitate");
            Cantitate.InnerText = line.Cantitate;
            XmlNode Pret = doc.CreateElement("Pret");
            Pret.InnerText = line.Pret;
            XmlNode Valoare = doc.CreateElement("Valoare");
            Valoare.InnerText = line.Valoare;

            XmlNode ProcTVA = doc.CreateElement("ProcTVA");
            ProcTVA.InnerText = line.ProcTVA;
            XmlNode TVA = doc.CreateElement("TVA");
            TVA.InnerText = line.TVA;
            XmlNode Cont = doc.CreateElement("Cont");
            Cont.InnerText = line.Cont;

            linie.AppendChild(LinieNrCrt);
            linie.AppendChild(Descriere);
            linie.AppendChild(CodArticolFurnizor);
            linie.AppendChild(CodArticolClient);
            linie.AppendChild(CodBare);

            linie.AppendChild(InformatiiSuplimentare);
            linie.AppendChild(UM);
            linie.AppendChild(Cantitate);
            linie.AppendChild(Pret);
            linie.AppendChild(Valoare);

            linie.AppendChild(ProcTVA);
            linie.AppendChild(TVA);
            linie.AppendChild(Cont);

            #endregion

            doc.Save(xmlFileName + ".xml");  
        }
    }
}
