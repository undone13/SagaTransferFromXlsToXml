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
using System.Threading;
using System.Linq.Expressions;
using static SagaTransferFromXlsToXml.InvoiceClass;

namespace SagaTransferFromXlsToXml
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string fileName = "documentXML";
        public string filePath = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonChooseFile_Click(object sender, RoutedEventArgs e)
        {
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

            buttonChooseFile.Visibility = Visibility.Collapsed;
            imageLoading.Visibility = Visibility.Visible;

            

            if(File.Exists(fileName) == false)
            {
                CreateXML(fileName);
            }
            else
            {
                //delete file
                File.Delete(fileName + ".xml");
            }

            Thread thread = new Thread(ReadAndWrite);
            thread.Start();
        }

        private void ReadAndWrite()
        {
            List<string> headers = new List<string>();

            try
            {
                int defaultAccount = filePath.ToLower().Contains("intrare") ? 371 : 707;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {                    
                        //do
                        //{
                        //    while (reader.Read())
                        //    {
                        //        InvoiceClass.Antet antet = new InvoiceClass.Antet();
                        //        antet.ClientNume = reader.GetString(0);
                        //        antet.ClientCIF = reader.GetString(1);
                        //        antet.FacturaNumar = reader.GetString(2);  //reader.GetString(2).Split(' ').Skip(1).FirstOrDefault(); //fact 1234 -> 1234
                        //        antet.FurnizorNume = reader.GetString(3);
                        //        antet.FurnizorCIF = reader.GetString(4);
                        //        antet.FacturaData = reader.GetString(5);
                        //        antet.FacturaMoneda = reader.GetString(9);

                        //        InvoiceClass.Linie line = new InvoiceClass.Linie();
                        //        line.LinieNrCrt = "1";
                        //        line.Descriere = "Marfa";
                        //        line.Cantitate = "1";
                        //        line.Pret = reader.GetDouble(8).ToString();
                        //        line.Valoare = line.Pret;
                        //        double TVAproc = reader.GetDouble(7) / reader.GetDouble(6) * 100f;
                        //        line.ProcTVA = Convert.ToInt32(TVAproc).ToString();
                        //        line.TVA = reader.GetDouble(7).ToString();
                        //        line.Cont = "371";

                        //        CreateInvoice(fileName, antet, line);
                        //    }
                        //} while (reader.NextResult());

                        var result = reader.AsDataSet();
                        var rowsCount = result.Tables[0].Rows.Count;
                        var columnsCount = result.Tables[0].Columns.Count;

                        for(int i=1; i< rowsCount; i++)
                        {
                            InvoiceClass.Antet antet = new InvoiceClass.Antet();
                            InvoiceClass.Linie line = new InvoiceClass.Linie();
                            List<InvoiceClass.Linie> lines = new List<InvoiceClass.Linie>();
                            for (int j = 0; j < columnsCount; j++)
                            {
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorNume")
                                {
                                    antet.FurnizorNume = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorCIF")
                                {
                                    antet.FurnizorCIF = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorNrRegCom")
                                {
                                    antet.FurnizorNrRegCom = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorCapital")
                                {
                                    antet.FurnizorCapital = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorTara")
                                {
                                    antet.FurnizorTara = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorJudet")
                                {
                                    antet.FurnizorJudet = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorAdresa")
                                {
                                    antet.FurnizorAdresa = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorBanca")
                                {
                                    antet.FurnizorBanca = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorIBAN")
                                {
                                    antet.FurnizorIBAN = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FurnizorInformatiiSuplimentare")
                                {
                                    antet.FurnizorInformatiiSuplimentare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientNume")
                                {
                                    antet.ClientNume = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientInformatiiSuplimentare")
                                {
                                    antet.ClientInformatiiSuplimentare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientCIF")
                                {
                                    antet.ClientCIF = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientNrRegCom")
                                {
                                    antet.ClientNrRegCom = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientTara")
                                {
                                    antet.ClientTara = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientJudet")
                                {
                                    antet.ClientJudet = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientAdresa")
                                {
                                    antet.ClientAdresa = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientBanca")
                                {
                                    antet.ClientBanca = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientIBAN")
                                {
                                    antet.ClientIBAN = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientTelefon")
                                {
                                    antet.ClientTelefon = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "ClientMail")
                                {
                                    antet.ClientMail = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaNumar")
                                {
                                    antet.FacturaNumar = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaData")
                                {
                                    antet.FacturaData = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaScadenta")
                                {
                                    antet.FacturaScadenta = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaTaxareInversa")
                                {
                                    antet.FacturaTaxareInversa = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaTVAIncasare")
                                {
                                    antet.FacturaTVAIncasare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaTip")
                                {
                                    antet.FacturaTip = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaInformatiiSuplimentare")
                                {
                                    antet.FacturaInformatiiSuplimentare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaInformatiiSuplimentare")
                                {
                                    antet.FacturaInformatiiSuplimentare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaMoneda")
                                {
                                    antet.FacturaMoneda = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaCotaTVA")
                                {
                                    antet.FacturaCotaTVA = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaID")
                                {
                                    antet.FacturaID = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "FacturaGreutate")
                                {
                                    antet.FacturaGreutate = result.Tables[0].Rows[i][j].ToString();
                                }

                                //Linie
                                if (result.Tables[0].Rows[0][j].ToString() == "LinieNrCrt")
                                {
                                    line.LinieNrCrt = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Gestiune")
                                {
                                    line.Gestiune = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Activitate")
                                {
                                    line.Activitate = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Descriere")
                                {
                                    line.Descriere = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "CodArticolFurnizor")
                                {
                                    line.CodArticolFurnizor = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "CodArticolClient")
                                {
                                    line.CodArticolClient = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "CodBare")
                                {
                                    line.CodBare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "InformatiiSuplimentare")
                                {
                                    line.InformatiiSuplimentare = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "UM")
                                {
                                    line.UM = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Cantitate")
                                {
                                    line.Cantitate = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Pret")
                                {
                                    line.Pret = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "Valoare")
                                {
                                    line.Valoare = result.Tables[0].Rows[i][j].ToString();
                                }
                                //if (result.Tables[0].Rows[0][j].ToString() == "CotaTVA")
                                //{
                                //    line.CotaTVA = result.Tables[0].Rows[i][j].ToString();
                                //}
                                if (result.Tables[0].Rows[0][j].ToString() == "ProcTVA")
                                {
                                    line.ProcTVA = result.Tables[0].Rows[i][j].ToString();
                                }
                                if (result.Tables[0].Rows[0][j].ToString() == "TVA")
                                {
                                    line.TVA = result.Tables[0].Rows[i][j].ToString();
                                }
                            }
                            if(string.IsNullOrEmpty(line.Cont))
                            {
                                line.Cont = defaultAccount.ToString();
                            }
                            if(string.IsNullOrEmpty(line.Descriere))
                            {
                                line.Descriere = "marfa";
                            }
                            if(string.IsNullOrEmpty(line.UM))
                            {
                                line.UM = "buc";
                            }
                            //if(string.IsNullOrEmpty(antet.FacturaCotaTVA)) 
                            //{
                            //    var TVAvalue = Convert.ToDouble(line.TVA) / Convert.ToDouble(line.Pret) * 100.00f;
                            //    antet.FacturaCotaTVA = Convert.ToInt32(TVAvalue).ToString();
                            //}
                            if(string.IsNullOrEmpty(line.Cantitate))
                            {
                                line.Cantitate = "1";
                            }
                            if(string.IsNullOrEmpty(line.Valoare))
                            {
                                var val = Convert.ToDouble(line.Cantitate) * Convert.ToDouble(line.Pret);
                                line.Valoare = val.ToString();
                            }
                            //if(string.IsNullOrEmpty(line.CotaTVA))
                            //{
                            //    var TVAvalue = Convert.ToDouble(line.TVA) / Convert.ToDouble(line.Pret) * 100.00f;
                            //    line.CotaTVA = Convert.ToInt32(TVAvalue).ToString();
                            //}
                            if (string.IsNullOrEmpty(line.ProcTVA))
                            {
                                var TVAvalue = Convert.ToDouble(line.TVA) / Convert.ToDouble(line.Pret) * 100.00f;
                                line.ProcTVA = Convert.ToInt32(TVAvalue).ToString();
                            }
                            if(string.IsNullOrEmpty(antet.ClientTara))
                            {
                                antet.ClientTara = "RO";
                            }
                            if (string.IsNullOrEmpty(antet.FurnizorTara))
                            {
                                antet.FurnizorTara = "RO";
                            }
                            if(string.IsNullOrEmpty(antet.ClientCIF))
                            {
                                antet.ClientCIF = "-";
                            }
                            if(string.IsNullOrEmpty(antet.FacturaMoneda))
                            {
                                antet.FacturaMoneda = "RON";
                            }
                            if(string.IsNullOrEmpty(antet.FacturaGreutate))
                            {
                                antet.FacturaGreutate = "0.000";
                            }
                            if(string.IsNullOrEmpty(line.LinieNrCrt))
                            {
                                line.LinieNrCrt = "1";
                            }

                            lines.Add(line);
                            CreateInvoice(fileName, antet, lines);
                        }

                    }
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    buttonChooseFile.Visibility = Visibility.Visible;
                    imageLoading.Visibility = Visibility.Collapsed;
                }));

                MessageBox.Show("Fisierul s-a creat cu succes!");
                System.Diagnostics.Process.Start(fileName + ".xml");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void CreateXML(string xmlFileName) 
        {
            XmlDocument doc = new XmlDocument();

            XmlNode facturi = doc.CreateElement("Facturi");
            doc.AppendChild(facturi);

            doc.Save(xmlFileName + ".xml");
        }

        private void CreateInvoice(string xmlFileName, InvoiceClass.Antet antet, List<InvoiceClass.Linie> lines)
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
            XmlNode FurnizorTara = doc.CreateElement("FurnizorTara");
            FurnizorTara.InnerText = antet.FurnizorTara;
            XmlNode FurnizorJudet = doc.CreateElement("FurnizorJudet");
            FurnizorJudet.InnerText = antet.FurnizorJudet;
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
            XmlNode ClientTara = doc.CreateElement("ClientTara");
            ClientTara.InnerText = antet.ClientTara;
            XmlNode ClientJudet = doc.CreateElement("ClientJudet");
            ClientJudet.InnerText = antet.ClientJudet;
            XmlNode ClientAdresa = doc.CreateElement("ClientAdresa");
            ClientAdresa.InnerText = antet.ClientAdresa;
            XmlNode ClientBanca = doc.CreateElement("ClientBanca");
            ClientBanca.InnerText = antet.ClientBanca;

            XmlNode ClientIBAN = doc.CreateElement("ClientIBAN");
            ClientIBAN.InnerText = antet.ClientIBAN;
            XmlNode ClientTelefon = doc.CreateElement("ClientTelefon");
            ClientTelefon.InnerText = antet.ClientTelefon;
            XmlNode ClientMail = doc.CreateElement("ClientMail");
            ClientMail.InnerText = antet.ClientMail;
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
            XmlNode FacturaTip = doc.CreateElement("FacturaTip");
            FacturaTip.InnerText = antet.FacturaTip;
            XmlNode FacturaInformatiiSuplimentare = doc.CreateElement("FacturaInformatiiSuplimentare");
            FacturaInformatiiSuplimentare.InnerText = antet.FacturaInformatiiSuplimentare;
            XmlNode FacturaMoneda = doc.CreateElement("FacturaMoneda");
            FacturaMoneda.InnerText = antet.FacturaMoneda;
            XmlNode FacturaCotaTVA = doc.CreateElement("FacturaCotaTVA");
            FacturaCotaTVA.InnerText = antet.FacturaCotaTVA;
            XmlNode FacturaID = doc.CreateElement("FacturaID");
            FacturaID.InnerText = antet.FacturaID;
            XmlNode FacturaGreutate = doc.CreateElement("FacturaGreutate");
            FacturaGreutate.InnerText = antet.FacturaGreutate;
            

            antetNode.AppendChild(FurnizorNume);
            antetNode.AppendChild(FurnizorCIF);
            antetNode.AppendChild(FurnizorNrRegCom);
            antetNode.AppendChild(FurnizorCapital);
            antetNode.AppendChild(FurnizorTara); //
            antetNode.AppendChild(FurnizorJudet); //
            antetNode.AppendChild(FurnizorAdresa);

            antetNode.AppendChild(FurnizorBanca);
            antetNode.AppendChild(FurnizorIBAN);
            antetNode.AppendChild(FurnizorInformatiiSuplimentare);
            antetNode.AppendChild(ClientNume);
            antetNode.AppendChild(ClientInformatiiSuplimentare);

            antetNode.AppendChild(ClientCIF);
            antetNode.AppendChild(ClientNrRegCom);
            antetNode.AppendChild(ClientTara); //
            antetNode.AppendChild(ClientJudet);
            antetNode.AppendChild(ClientAdresa);
            antetNode.AppendChild(ClientBanca);

            antetNode.AppendChild(ClientIBAN);
            antetNode.AppendChild(ClientTelefon); // 
            antetNode.AppendChild(ClientMail); //
            antetNode.AppendChild(FacturaNumar);
            antetNode.AppendChild(FacturaData);
            antetNode.AppendChild(FacturaScadenta);
            antetNode.AppendChild(FacturaTaxareInversa);

            antetNode.AppendChild(FacturaTVAIncasare);
            antetNode.AppendChild(FacturaTip); // 
            antetNode.AppendChild(FacturaInformatiiSuplimentare);
            antetNode.AppendChild(FacturaMoneda);
            antetNode.AppendChild(FacturaCotaTVA); //
            antetNode.AppendChild(FacturaID); //
            antetNode.AppendChild(FacturaGreutate); //
            #endregion

            #region Linii
            foreach(var line in lines)
            {
                XmlNode detalii = doc.CreateElement("Detalii");
                factura.AppendChild(detalii);
                XmlNode continut = doc.CreateElement("Continut");
                detalii.AppendChild(continut);
                XmlNode linie = doc.CreateElement("Linie");
                continut.AppendChild(linie);

                XmlNode LinieNrCrt = doc.CreateElement("LinieNrCrt");
                LinieNrCrt.InnerText = line.LinieNrCrt;
                XmlNode Gestiune = doc.CreateElement("Gestiune");
                Gestiune.InnerText = line.Gestiune;
                XmlNode Activitate = doc.CreateElement("Activitate");
                Activitate.InnerText = line.Activitate;

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


                //XmlNode CotaTVA = doc.CreateElement("CotaTVA");
                //CotaTVA.InnerText = line.CotaTVA;
                XmlNode ProcTVA = doc.CreateElement("ProcTVA");
                ProcTVA.InnerText = line.ProcTVA;
                XmlNode TVA = doc.CreateElement("TVA");
                TVA.InnerText = line.TVA;
                XmlNode Cont = doc.CreateElement("Cont");
                Cont.InnerText = line.Cont;

                linie.AppendChild(LinieNrCrt);
                linie.AppendChild(Gestiune); //
                linie.AppendChild(Activitate); //
                linie.AppendChild(Descriere);
                linie.AppendChild(CodArticolFurnizor);
                linie.AppendChild(CodArticolClient);
                linie.AppendChild(CodBare);

                linie.AppendChild(InformatiiSuplimentare);
                linie.AppendChild(UM);
                linie.AppendChild(Cantitate);
                linie.AppendChild(Pret);
                linie.AppendChild(Valoare);

                //linie.AppendChild(CotaTVA); //
                linie.AppendChild(ProcTVA);
                linie.AppendChild(TVA);
                linie.AppendChild(Cont);
            }
            #endregion

            doc.Save(xmlFileName + ".xml");  
        }
    }
}
