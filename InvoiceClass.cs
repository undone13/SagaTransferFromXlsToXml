﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SagaTransferFromXlsToXml
{
    internal class InvoiceClass
    {
        public class Antet
        {
            public string FurnizorNume;
            public string FurnizorCIF;
            public string FurnizorNrRegCom;
            public string FurnizorCapital;
            public string FurnizorAdresa;
            public string FurnizorBanca;
            public string FurnizorIBAN;
            public string FurnizorInformatiiSuplimentare;
            public string ClientNume;
            public string ClientInformatiiSuplimentare;
            public string ClientCIF;
            public string ClientNrRegCom;
            public string ClientJudet;
            public string ClientAdresa;
            public string ClientBanca;
            public string ClientIBAN;
            public string FacturaNumar;
            public string FacturaData;
            public string FacturaScadenta;
            public string FacturaTaxareInversa;
            public string FacturaTVAIncasare;
            public string FacturaInformatiiSuplimentare;
            public string FacturaMoneda;
        }
        public class Linie
        {
            public string LinieNrCrt;
            public string Descriere;
            public string CodArticolFurnizor;
            public string CodArticolClient;
            public string CodBare;
            public string InformatiiSuplimentare;
            public string UM;
            public string Cantitate;
            public string Pret;
            public string Valoare;
            public string ProcTVA;
            public string TVA;
            public string Cont;
        }
    }
}