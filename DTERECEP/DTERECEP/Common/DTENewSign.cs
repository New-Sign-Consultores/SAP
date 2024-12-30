using DTE;
using DTERECEP.Common.CRUD;
using DTERECEP.Common.DTE;
using DTERECEP.Common.NEWSIGN;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace DTERECEP.Common
{
    public class DTENewSign
    {
        CultureInfo cultureInfo = new CultureInfo("es-CL");
        string apiUrl = "http://172.16.0.13/api/Core.svc/core";
        string apiKey = "2ec64694-829f-4a6e-b369-dc8936320e09";
        string query = "(FchEmis:2023-08-08)";
        string Environment = "P";
        public void GETListDTE(string FchDesde, string FchHasta, string TipoDoc)
        {
            string Ffecha = null,FTipo = null,RUTRecep =null;
            RUTRecep = "78058870-5";
            Common common = new Common();
            try
            {
                DateTime FchDesd, FchHast;
                if (string.IsNullOrEmpty(FchDesde) && string.IsNullOrEmpty(FchHasta))
                {
                    FchDesd = DateTime.Now.AddDays(-8);
                    FchHast = DateTime.Now;
                }
                else
                {
                    FchDesd = Convert.ToDateTime(FchDesde, cultureInfo);
                    FchHast = Convert.ToDateTime(FchHasta, cultureInfo);
                }
                switch(TipoDoc)
                {
                    case "33":
                        FTipo = " AND (TipoDTE:33))";
                        break;
                    case "34":
                        FTipo = " AND (TipoDTE:34))";
                        break;
                    case "56":
                        FTipo = " AND (TipoDTE:56))";
                        break;
                    case "61":
                        FTipo = " AND (TipoDTE:61))";
                        break;
                    case "Todos":
                        FTipo = " AND (TipoDTE:33 OR TipoDTE:34 OR TipoDTE:56 OR TipoDTE:61))";
                        break;
                }
                RUTRecep = "( (RUTRecep:" + RUTRecep + ") AND ";
                Ffecha = " FchEmis:["+ FchDesd.ToString("yyyy-MM-dd") + " TO " + FchHast.ToString("yyyy-MM-dd") + "]";
                query = RUTRecep + Ffecha + FTipo;
                string QueryBase64 = Base64Encode(query);
                string Objeto = "/PaginatedSearch/"+ Environment +"/R/" + QueryBase64 +"/1/500";
                var restClient = new RestClient();
                var restRequest = new RestRequest(apiUrl + Objeto, Method.GET);
                restRequest.AddHeader("AuthKey", apiKey);
                restRequest.AddHeader("Content-Type", "application/json");

                int CantDocSAP = 0;
                IRestResponse restResponse = restClient.Execute(restRequest);
                if (restResponse.StatusDescription.Equals("OK"))
                {
                    SearchResultXML searchResultXML = GetSearchResultXML(restResponse.Content);
                    ResultSt resultSt = new ResultSt();
                    CantDocSAP = common.DocInTable(FchDesde, FchHasta, TipoDoc);

                    if (searchResultXML.TotalDocuments > 0)
                    {
                        if (searchResultXML.TotalDocuments != CantDocSAP)
                        {
                            foreach (DocumentXML documentXML in searchResultXML.Data.ltDocuments)
                            {
                                //string Tipo = documentXML.TipoDTE;
                                //string Folio = documentXML.Folio;
                                //string RUT = documentXML.RUTEmisor;
                                //string code = Tipo.PadLeft(5,'0') + Folio.PadLeft(50,'0')+RUT.PadLeft(12,'0');
                                if (ValidarDTE(documentXML.RUTEmisor, documentXML.TipoDTE, documentXML.Folio))
                                {

                                }
                                else
                                {
                                    CRUD_ASRDTE cRUD_ASRDTE = new CRUD_ASRDTE();
                                    ASRDTE aSRDTE = new ASRDTE();
                                    if (documentXML.TieneArchivo)
                                    {
                                        ResultDTE resultDTE = new ResultDTE();
                                        resultDTE = GETXMLDTE(documentXML);
                                        //GETPDFDTE(documentXML);

                                        if (resultDTE.Success)
                                        {            
                                            //documentXML.Statuses
                                            aSRDTE.DocumentID = documentXML.DocumentID;
                                            aSRDTE.RutEmisor = documentXML.RUTEmisor;
                                            aSRDTE.RznSoc = resultDTE.DTE.Emisor.RznSoc;
                                            aSRDTE.TipoDTE = documentXML.TipoDTE;
                                            aSRDTE.ExiEmisor = false;
                                            aSRDTE.Folio = documentXML.Folio;
                                            aSRDTE.FchEmis = Convert.ToDateTime(resultDTE.DTE.IdDoc.FchEmis);
                                            try
                                            {
                                                aSRDTE.FchVenc = Convert.ToDateTime(resultDTE.DTE.IdDoc.FchVenc,System.Globalization.CultureInfo.InvariantCulture);
                                            }
                                            catch { }
                                            aSRDTE.FmaPago = Convert.ToString(resultDTE.DTE.IdDoc.FmaPago);
                                            aSRDTE.MntNeto = documentXML.MntNeto;
                                            aSRDTE.MntExe = resultDTE.DTE.Totales.MntExe;
                                            aSRDTE.TasaIVA = resultDTE.DTE.Totales.TasaIVA;
                                            aSRDTE.IVA = resultDTE.DTE.Totales.IVA;
                                            aSRDTE.MntTotal = resultDTE.DTE.Totales.MntTotal;
                                            //Validar
                                            string FoliosOC = null, FoliosEM = null, FoliosFA = null;
                                            if (resultDTE.DTE.Referencia.Count > 0)
                                            {

                                                #region Referencias del XML
                                                var RefOC = new List<Referencia>();
                                                var RefEM = new List<Referencia>();
                                                var RefFA = new List<Referencia>();

                                                RefOC = resultDTE.DTE.Referencia.Where(j => j.TpoDocRef == "801").ToList();//|| j.TpoDocRef == "802").ToList();
                                                RefEM = resultDTE.DTE.Referencia.Where(j => j.TpoDocRef == "52" || j.TpoDocRef == "HES" || j.TpoDocRef == "50").ToList();
                                                RefFA = resultDTE.DTE.Referencia.Where(j => j.TpoDocRef == "33" || j.TpoDocRef == "34").ToList();
                                                FoliosOC = string.Join(", ", RefOC.Select(z => z.FolioRef));
                                                FoliosEM = string.Join(", ", RefEM.Select(z => z.FolioRef));
                                                FoliosFA = string.Join(", ", RefFA.Select(z => z.FolioRef));
                                                #endregion Referencias del XML
                                            }
                                            aSRDTE.FolioRefOC = (string.IsNullOrEmpty(FoliosOC) ? string.Empty : FoliosOC);
                                            aSRDTE.FolioSAPOC = string.Empty;
                                            aSRDTE.FolioRefEM = (string.IsNullOrEmpty(FoliosEM) ? string.Empty : FoliosEM);
                                            aSRDTE.FolioSAPEM = string.Empty;
                                            aSRDTE.FolioRefFA = (string.IsNullOrEmpty(FoliosFA) ? string.Empty : FoliosFA);
                                            aSRDTE.FolioSAPFA = string.Empty;

                                            #region Valida si existe en SAP
                                            resultSt = cRUD_ASRDTE.ExistSAP(documentXML.RUTEmisor, documentXML.TipoDTE, documentXML.Folio);
                                            if (resultSt.Estado)
                                            {
                                                aSRDTE.DocEntryS = resultSt.Codigo;
                                                aSRDTE.ObjType = resultSt.Valor;
                                            }
                                            else
                                            {
                                                aSRDTE.DocEntryS = string.Empty;
                                                aSRDTE.ObjType = string.Empty;
                                            }
                                            #endregion Valida si existe en SAP

                                            aSRDTE.XML = resultDTE.XMLString;
                                            aSRDTE.PDF64 = string.Empty;

                                            cRUD_ASRDTE.AddASRDTE(aSRDTE);
                                        }
                                    }
                                    else //No tiene archivo
                                    {
                                        aSRDTE.DocumentID = documentXML.DocumentID;
                                        aSRDTE.RutEmisor = documentXML.RUTEmisor;
                                        aSRDTE.RznSoc = documentXML.RznSoc;
                                        aSRDTE.TipoDTE = documentXML.TipoDTE;
                                        aSRDTE.ExiEmisor = false;
                                        aSRDTE.Folio = documentXML.Folio;
                                        aSRDTE.FchEmis = Convert.ToDateTime(documentXML.FchEmis);
                                        //aSRDTE.FchVenc = Convert.ToDateTime(resultDTE.DTE.IdDoc.FchVenc);
                                        aSRDTE.FmaPago = Convert.ToString(documentXML.FmaPago);
                                        aSRDTE.MntNeto = documentXML.MntNeto;
                                        aSRDTE.MntExe = 0;
                                        aSRDTE.TasaIVA = 0;
                                        aSRDTE.IVA = documentXML.IVA;
                                        aSRDTE.MntTotal = documentXML.MntTotal;
                                        //Validar
                                        aSRDTE.FolioRefOC = string.Empty;
                                        aSRDTE.FolioSAPOC = string.Empty;
                                        aSRDTE.FolioRefEM = string.Empty;
                                        aSRDTE.FolioSAPEM = string.Empty;
                                        aSRDTE.FolioRefFA = string.Empty;
                                        aSRDTE.FolioSAPFA = string.Empty;

                                        #region Valida si existe en SAP
                                        resultSt = cRUD_ASRDTE.ExistSAP(documentXML.RUTEmisor, documentXML.TipoDTE, documentXML.Folio);
                                        if (resultSt.Estado)
                                        {
                                            aSRDTE.DocEntryS = resultSt.Codigo;
                                            aSRDTE.ObjType = resultSt.Valor;
                                        }
                                        else
                                        {
                                            aSRDTE.DocEntryS = string.Empty;
                                            aSRDTE.ObjType = string.Empty;
                                        }
                                        #endregion Valida si existe en SAP

                                        aSRDTE.XML = string.Empty;
                                        aSRDTE.PDF64 = string.Empty;

                                        cRUD_ASRDTE.AddASRDTE(aSRDTE);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public bool ValidarDTE(string RUT, string Tipo, string Folio)
        {
            bool existe = false;
            CRUD_ASRDTE cRUD_ASRDTE = new CRUD_ASRDTE();
            string Code = cRUD_ASRDTE.ExistASRDTE(RUT, Tipo, Folio);
            if (!string.IsNullOrEmpty(Code))
                existe = true;
            return existe;
         }

        public ResultDTE GETXMLDTE(DocumentXML documentXML)
        {
            ResultDTE resultDTE = new ResultDTE();
            try
            {   
                string Objeto = "/RecoverXML_V2";
                var restClient = new RestClient();
                var restRequest = new RestRequest(apiUrl + Objeto, Method.POST);
                restRequest.AddHeader("AuthKey", apiKey);
                restRequest.AddHeader("Content-Type", "application/json");
                var body = @"{                                                  " + "\n" +
                @"	""Environment"":""" + Environment +@""",                    " + "\n" +
                @"	""Group"":""R"",                                            " + "\n" +
                @"	""Rut"":""" + documentXML.RUTEmisor + @""",                 " + "\n" +
                @"	""DocType"":""" + documentXML.TipoDTE + @""",               " + "\n" +
                @"	""Folio"":""" + documentXML.Folio + @""",                   " + "\n" +
                @"	""IsForDistribution"":""true""          " + "\n" +
                @"}";
                //restRequest.AddBody(body);
                restRequest.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse restResponse = restClient.Execute(restRequest);

                if (restResponse.StatusDescription.Equals("OK"))
                {
                    RecoverXML_V2XML recoverXML_V2XML = GetRecoverXML_V2(restResponse.Content);
                    if (recoverXML_V2XML.Result == "0")
                    {
                        resultDTE = ObtenerDTE(recoverXML_V2XML.Data);                       
                    }
                    else
                    {
                        resultDTE.Success = false;
                        //error
                    }
                }
                else
                {
                    resultDTE.Success = false; 
                    resultDTE.Mensaje = "Not Found";
                }
            }
            catch(Exception ex)
            {
                resultDTE.Success = false;
                resultDTE.Mensaje = ex.Message;
            }
            return resultDTE;
        }

        public DTECompany getConfig()
        {
            DTECompany dTECompany = new DTECompany();
            string Query = @" SELECT 
	                         T0.""U_ENVIRONMENT""
	                        ,T0.""U_RUTREC""
	                        ,T0.""U_SENTACD""
	                        ,T0.""U_SENTRZD""
	                        ,T0.""U_URLDTELIST""
	                        ,T0.""U_URLDTE""
	                        ,T0.""U_URLACD""
	                        ,T0.""U_URLRZD""
	                        ,T0.""U_USER""
	                        ,T0.""U_KEY""
                            ,T0.""U_LKPRCOR""
                            ,T0.""U_LKPRCDN""
                        FROM ""@ASCFRC"" T0 ";

            try
            {
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                recordset.DoQuery(Query);
                if (!recordset.EoF)
                {
                    dTECompany.ENVIRONMENT = Convert.ToString(recordset.Fields.Item("U_ENVIRONMENT").Value);
                    dTECompany.RUTREC = Convert.ToString(recordset.Fields.Item("U_RUTREC").Value);
                    dTECompany.URLDTELIST = Convert.ToString(recordset.Fields.Item("U_URLDTELIST").Value);
                    dTECompany.URLDTE = Convert.ToString(recordset.Fields.Item("U_URLDTE").Value);
                    dTECompany.KEY = Convert.ToString(recordset.Fields.Item("U_KEY").Value);
                    dTECompany.LKPRCOR = Convert.ToString(recordset.Fields.Item("U_LKPRCOR").Value);
                    dTECompany.LKPRCDN = Convert.ToString(recordset.Fields.Item("U_LKPRCDN").Value);
                }

            }
            catch { }
            return dTECompany;
        }
        public String GETPDFDTE(string RUTEmisor, string TipoDTE, string Folio)
        {
            String Base64 = null;
            DTECompany dTECompany = new DTECompany();
            dTECompany = getConfig();
            try
            {
                string Objeto = "/RecoverPDF_V2";
                #region Filtros 
                string Filter = "/"+ dTECompany.ENVIRONMENT + "/R/" + RUTEmisor +"/" + TipoDTE +"/"+ Folio;
                Objeto = Objeto + Filter;
                #endregion Filtros
                var restClient = new RestClient();
                var restRequest = new RestRequest(dTECompany.URLDTE + Objeto, Method.GET);
                restRequest.AddHeader("AuthKey", dTECompany.KEY );
                //restRequest.AddHeader("Content-Type", "application/json");
                //var body = @"{  }";
                //restRequest.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse restResponse = restClient.Execute(restRequest);
                if (restResponse.StatusDescription.Equals("OK"))
                {
                    RecoverXML_V2XML RecoverPDF_V2 = GetRecoverPDF_V2(restResponse.Content);
                    if (RecoverPDF_V2.Result == "0")
                    {
                        Base64 = RecoverPDF_V2.Data;
                    }
                    else
                    {
                        //error
                    }
                }
                else
                {
                    //Not Found
                }
            }
            catch// (Exception ex)
            {

            }
            return Base64;
        }

        private RecoverXML_V2XML GetRecoverXML_V2(string RecoverXML_V2String)
        {
            RecoverXML_V2XML recoverXML_V2XML = new RecoverXML_V2XML();
            //string StringToXml = Base64Decode(searchResultString);
            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(RecoverXML_V2String);
                //string FD = xmlSearchResul.ChildNodes.Item(0);// new System.Linq.SystemCore_EnumerableDebugView(xmlSearchResul.ChildNodes.Item(0)).Items[0];
                //XmlNamespaceManager ns = new XmlNamespaceManager(xmlSearchResul.NameTable);
                //String PathSearchResult = "SearchResult";
                //XmlNode SearchResul = xmlSearchResul.SelectSingleNode(PathSearchResult, ns);
                XmlNodeList nodeList = xmlDocument.DocumentElement.ChildNodes;
                foreach (XmlNode xmlNode in nodeList)
                {
                    switch (xmlNode.Name)
                    {
                        case "Data":
                            string DataDecode = Base64Decode(xmlNode.InnerText);                                               
                            recoverXML_V2XML.Data = DataDecode;
                            break;
                        case "Description":
                            recoverXML_V2XML.Description = xmlNode.InnerText;
                            break;
                        case "Result":
                            recoverXML_V2XML.Result = xmlNode.InnerText;
                            break;
                        case "StackTrace":
                            recoverXML_V2XML.StackTrace = xmlNode.InnerText;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                recoverXML_V2XML.Result = "-1";
                recoverXML_V2XML.Description = ex.Message;
            }
            return recoverXML_V2XML;
        }

        private RecoverXML_V2XML GetRecoverPDF_V2(string RecoverPDF_V2String)
        {
            RecoverXML_V2XML recoverXML_V2XML = new RecoverXML_V2XML();
            //string StringToXml = Base64Decode(searchResultString);
            try
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(RecoverPDF_V2String);
                //string FD = xmlSearchResul.ChildNodes.Item(0);// new System.Linq.SystemCore_EnumerableDebugView(xmlSearchResul.ChildNodes.Item(0)).Items[0];
                //XmlNamespaceManager ns = new XmlNamespaceManager(xmlSearchResul.NameTable);
                //String PathSearchResult = "SearchResult";
                //XmlNode SearchResul = xmlSearchResul.SelectSingleNode(PathSearchResult, ns);
                XmlNodeList nodeList = xmlDocument.DocumentElement.ChildNodes;
                foreach (XmlNode xmlNode in nodeList)
                {
                    switch (xmlNode.Name)
                    {
                        case "Data":
                            string DataDecode = xmlNode.InnerText;
                            recoverXML_V2XML.Data = DataDecode;
                            break;
                        case "Description":
                            recoverXML_V2XML.Description = xmlNode.InnerText;
                            break;
                        case "Result":
                            recoverXML_V2XML.Result = xmlNode.InnerText;
                            break;
                        case "StackTrace":
                            recoverXML_V2XML.StackTrace = xmlNode.InnerText;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                recoverXML_V2XML.Result = "-1";
                recoverXML_V2XML.Description = ex.Message;
            }
            return recoverXML_V2XML;
        }

        public ResultDTE ObtenerDTE(String decodeStringXml)
        {
            ResultDTE resultDTE = new ResultDTE();
            DTERECEP.DTE.DTE objDTE = new DTERECEP.DTE.DTE();
            try
            {
                // crear documento xml para obtener datos
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(decodeStringXml);

                // namespace del S.I.I.
                //System.Xml.Linq.XNamespace ns = "http://www.sii.cl/SiiDte";

                //System.Xml.Linq.XElement documento = xmlDoc.Descendants(ns + "Documento").FirstOrDefault();
                XmlNamespaceManager ns = new XmlNamespaceManager(xmlDoc.NameTable);
                ns.AddNamespace("ns", "http://www.sii.cl/SiiDte");


                // NODO IDENTIFICACION DEL DOCUMENTO
                #region IDDOC

                //String PathIdDoc = "sii:DTE/sii:Documento/siï:Encabezado/sii:IdDoc";
                String PathIdDoc = "//ns:IdDoc";
                XmlNode IdentificacionDoc = xmlDoc.SelectSingleNode(PathIdDoc, ns);

                foreach (XmlNode childNode in IdentificacionDoc)
                {
                    switch (childNode.Name)
                    {
                        case "TipoDTE":
                            objDTE.IdDoc.TipoDTE = childNode.InnerText;
                            break;
                        case "Folio":
                            objDTE.IdDoc.Folio = Int64.Parse(childNode.InnerText);
                            break;
                        case "FchEmis":
                            objDTE.IdDoc.FchEmis = childNode.InnerText;
                            break;
                        case "IndNoRebaja":
                            objDTE.IdDoc.IndNoRebaja = Int32.Parse(childNode.InnerText);
                            break;
                        case "TipoDespacho":
                            objDTE.IdDoc.TipoDespacho = Int32.Parse(childNode.InnerText);
                            break;
                        case "IndTraslado":
                            objDTE.IdDoc.IndTraslado = Int32.Parse(childNode.InnerText);
                            break;
                        case "TpoImpresion":
                            objDTE.IdDoc.TpoImpresion = childNode.InnerText;
                            break;
                        case "IndServicio":
                            objDTE.IdDoc.IndServicio = Int32.Parse(childNode.InnerText);
                            break;
                        case "MntBruto":
                            objDTE.IdDoc.MntBruto = Int32.Parse(childNode.InnerText);
                            break;
                        case "FmaPago":
                            objDTE.IdDoc.FmaPago = Int32.Parse(childNode.InnerText);
                            break;
                        case "FmaPagExp":
                            objDTE.IdDoc.FmaPagExp = Int32.Parse(childNode.InnerText);
                            break;
                        case "FchCancel":
                            objDTE.IdDoc.FchCancel = childNode.InnerText;
                            break;
                        case "MntCancel":
                            objDTE.IdDoc.MntCancel = Int64.Parse(childNode.InnerText);
                            break;
                        case "SaldoInsol":
                            objDTE.IdDoc.SaldoInsol = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntPagos":
                            MntPagos mnt = new MntPagos();
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("FchPago")) { mnt.FchPago = child.InnerText; }
                                else if (child.Name.Equals("MntPago")) { mnt.MntPago = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("GlosaPagos")) { mnt.GlosaPagos = child.InnerText; }
                            }
                            objDTE.IdDoc.MntPagos.Add(mnt);
                            break;
                        case "PeriodoDesde":
                            objDTE.IdDoc.PeriodoDesde = childNode.InnerText;
                            break;
                        case "PeriodoHasta":
                            objDTE.IdDoc.PeriodoHasta = childNode.InnerText;
                            break;
                        case "MedioPago":
                            objDTE.IdDoc.MedioPago = childNode.InnerText;
                            break;
                        case "TipoCtaPago":
                            objDTE.IdDoc.TipoCtaPago = childNode.InnerText;
                            break;
                        case "NumCtaPago":
                            objDTE.IdDoc.NumCtaPago = childNode.InnerText;
                            break;
                        case "BcoPago":
                            objDTE.IdDoc.BcoPago = childNode.InnerText;
                            break;
                        case "TermPagoCdg":
                            objDTE.IdDoc.TermPagoCdg = childNode.InnerText;
                            break;
                        case "TermPagoGlosa":
                            objDTE.IdDoc.TermPagoGlosa = childNode.InnerText;
                            break;
                        case "TermPagoDias":
                            objDTE.IdDoc.TermPagoDias = childNode.InnerText;
                            break;
                        case "FchVenc":
                            objDTE.IdDoc.FchVenc = childNode.InnerText;
                            break;
                    }
                }

                #endregion

                // NODO EMISOR DEL DOCUMENTO
                #region EMISOR

                String PathEmisor = "//ns:Emisor";
                XmlNode Emisor = xmlDoc.SelectSingleNode(PathEmisor, ns);

                foreach (XmlNode childNode in Emisor)
                {
                    switch (childNode.Name)
                    {
                        case "RUTEmisor":
                            objDTE.Emisor.RUTEmisor = childNode.InnerText;
                            break;
                        case "RznSoc":
                            objDTE.Emisor.RznSoc = childNode.InnerText;
                            break;
                        case "GiroEmis":
                            objDTE.Emisor.GiroEmis = childNode.InnerText;
                            break;
                        case "Telefono":
                            objDTE.Emisor.Telefono = childNode.InnerText;
                            break;
                        case "CorreoEmisor":
                            objDTE.Emisor.CorreoEmisor = childNode.InnerText;
                            break;
                        case "Acteco":
                            objDTE.Emisor.Acteco = childNode.InnerText;
                            break;
                        case "GuiaExport":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("CdgTraslado")) { objDTE.Emisor.CdgTraslado = Int32.Parse(child.InnerText); }
                                else if (child.Name.Equals("FolioAut")) { objDTE.Emisor.FolioAut = Int32.Parse(child.InnerText); }
                                else if (child.Name.Equals("FchAut")) { objDTE.Emisor.FchAut = child.InnerText; }
                            }
                            break;
                        case "Sucursal":
                            objDTE.Emisor.Sucursal = childNode.InnerText;
                            break;
                        case "CdgSIISucur":
                            objDTE.Emisor.CdgSIISucur = childNode.InnerText;
                            break;
                        case "DirOrigen":
                            objDTE.Emisor.DirOrigen = childNode.InnerText;
                            break;
                        case "CmnaOrigen":
                            objDTE.Emisor.CmnaOrigen = childNode.InnerText;
                            break;
                        case "CiudadOrigen":
                            objDTE.Emisor.CiudadOrigen = childNode.InnerText;
                            break;
                        case "CdgVendedor":
                            objDTE.Emisor.CdgVendedor = childNode.InnerText;
                            break;
                        case "IdAdicEmisor":
                            objDTE.Emisor.IdAdicEmisor = childNode.InnerText;
                            break;
                    }
                }

                #endregion

                // NODO RECEPTOR DEL DOCUMENTO
                #region RECEPTOR

                String PathReceptor = "//ns:Receptor";
                XmlNode Receptor = xmlDoc.SelectSingleNode(PathReceptor, ns);

                foreach (XmlNode childNode in Receptor)
                {
                    switch (childNode.Name)
                    {
                        case "RUTRecep":
                            objDTE.Receptor.RUTRecep = childNode.InnerText;
                            break;
                        case "CdgIntRecep":
                            objDTE.Receptor.CdgIntRecep = childNode.InnerText;
                            break;
                        case "RznSocRecep":
                            objDTE.Receptor.RznSocRecep = childNode.InnerText;
                            break;
                        case "Extranjero":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("NumId")) { objDTE.Receptor.NumId = child.InnerText; }
                                else if (child.Name.Equals("Nacionalidad")) { objDTE.Receptor.Nacionalidad = child.InnerText; }
                                else if (child.Name.Equals("IdAdicRecep")) { objDTE.Receptor.IdAdicRecep = child.InnerText; }
                            }
                            break;
                        case "GiroRecep":
                            objDTE.Receptor.GiroRecep = childNode.InnerText;
                            break;
                        case "Contacto":
                            objDTE.Receptor.Contacto = childNode.InnerText;
                            break;
                        case "CorreoRecep":
                            objDTE.Receptor.CorreoRecep = childNode.InnerText;
                            break;
                        case "DirRecep":
                            objDTE.Receptor.DirRecep = childNode.InnerText;
                            break;
                        case "CmnaRecep":
                            objDTE.Receptor.CmnaRecep = childNode.InnerText;
                            break;
                        case "CiudadRecep":
                            objDTE.Receptor.CiudadRecep = childNode.InnerText;
                            break;
                        case "DirPostal":
                            objDTE.Receptor.DirPostal = childNode.InnerText;
                            break;
                        case "CmnaPostal":
                            objDTE.Receptor.CmnaPostal = childNode.InnerText;
                            break;
                        case "CiudadPostal":
                            objDTE.Receptor.CiudadPostal = childNode.InnerText;
                            break;
                    }
                }


                #endregion

                // NODO TRANSPORTE DEL DOCUMENTO
                #region TRANSPORTE

                String PathTransporte = "//ns:Transporte";
                XmlNode Transporte = xmlDoc.SelectSingleNode(PathTransporte, ns);

                if (Transporte != null)
                {
                    foreach (XmlNode childNode in Transporte)
                    {
                        switch (childNode.Name)
                        {
                            case "Patente":
                                objDTE.Transporte.Patente = childNode.InnerText;
                                break;
                            case "RUTTrans":
                                objDTE.Transporte.RUTTrans = childNode.InnerText;
                                break;
                            case "Chofer":
                                foreach (XmlNode child in childNode.ChildNodes)
                                {
                                    if (child.Name.Equals("RUTChofer")) { objDTE.Transporte.RUTChofer = child.InnerText; }
                                    else if (child.Name.Equals("NombreChofer")) { objDTE.Transporte.NombreChofer = child.InnerText; }
                                }
                                break;
                            case "DirDest":
                                objDTE.Transporte.DirDest = childNode.InnerText;
                                break;
                            case "CmnaDest":
                                objDTE.Transporte.CmnaDest = childNode.InnerText;
                                break;
                            case "CiudadDest":
                                objDTE.Transporte.CiudadDest = childNode.InnerText;
                                break;
                            case "Aduana":
                                foreach (XmlNode childNodeAduana in childNode)
                                {
                                    switch (childNodeAduana.Name)
                                    {
                                        case "CodModVenta":
                                            objDTE.Transporte.Aduana.CodModVenta = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "CodClauVenta":
                                            objDTE.Transporte.Aduana.CodClauVenta = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotClauVenta":
                                            objDTE.Transporte.Aduana.TotClauVenta = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodViaTransp":
                                            objDTE.Transporte.Aduana.CodViaTransp = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "NombreTransp":
                                            objDTE.Transporte.Aduana.NombreTransp = childNodeAduana.InnerText;
                                            break;
                                        case "RUTCiaTransp":
                                            objDTE.Transporte.Aduana.RUTCiaTransp = childNodeAduana.InnerText;
                                            break;
                                        case "NomCiaTransp":
                                            objDTE.Transporte.Aduana.NomCiaTransp = childNodeAduana.InnerText;
                                            break;
                                        case "IdAdicTransp":
                                            objDTE.Transporte.Aduana.IdAdicTransp = childNodeAduana.InnerText;
                                            break;
                                        case "Booking":
                                            objDTE.Transporte.Aduana.Booking = childNodeAduana.InnerText;
                                            break;
                                        case "Operador":
                                            objDTE.Transporte.Aduana.Operador = childNodeAduana.InnerText;
                                            break;
                                        case "CodPtoEmbarque":
                                            objDTE.Transporte.Aduana.CodPtoEmbarque = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "IdAdicPtoEmb":
                                            objDTE.Transporte.Aduana.IdAdicPtoEmb = childNodeAduana.InnerText;
                                            break;
                                        case "CodPtoDesemb":
                                            objDTE.Transporte.Aduana.CodPtoDesemb = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "IdAdicPtoDesemb":
                                            objDTE.Transporte.Aduana.IdAdicPtoDesemb = childNodeAduana.InnerText;
                                            break;
                                        case "Tara":
                                            objDTE.Transporte.Aduana.Tara = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "CodUnidMedTara":
                                            objDTE.Transporte.Aduana.CodUnidMedTara = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "PesoBruto":
                                            objDTE.Transporte.Aduana.PesoBruto = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodUnidPesoBruto":
                                            objDTE.Transporte.Aduana.CodUnidPesoBruto = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "PesoNeto":
                                            objDTE.Transporte.Aduana.PesoNeto = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodUnidPesoNeto":
                                            objDTE.Transporte.Aduana.CodUnidPesoNeto = Int32.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotItems":
                                            objDTE.Transporte.Aduana.TotItems = Int64.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TotBultos":
                                            objDTE.Transporte.Aduana.TotBultos = Int64.Parse(childNodeAduana.InnerText);
                                            break;
                                        case "TipoBultos":
                                            TipoBultos tip = new TipoBultos();
                                            foreach (XmlNode child in childNodeAduana)
                                            {
                                                if (child.Name.Equals("CodTpoBultos")) { tip.CodTpoBultos = Int32.Parse(child.InnerText); }
                                                else if (child.Name.Equals("CantBultos")) { tip.CantBultos = Int64.Parse(child.InnerText); }
                                                else if (child.Name.Equals("Marcas")) { tip.Marcas = child.InnerText; }
                                                else if (child.Name.Equals("IdContainer")) { tip.IdContainer = child.InnerText; }
                                                else if (child.Name.Equals("Sello")) { tip.Sello = child.InnerText; }
                                                else if (child.Name.Equals("EmisorSello")) { tip.EmisorSello = child.InnerText; }
                                            }
                                            objDTE.Transporte.Aduana.TipoBultos.Add(tip);
                                            break;
                                        case "MntFlete":
                                            objDTE.Transporte.Aduana.MntFlete = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "MntSeguro":
                                            objDTE.Transporte.Aduana.MntSeguro = Double.Parse(childNodeAduana.InnerText.Replace(".", ","));
                                            break;
                                        case "CodPaisRecep":
                                            objDTE.Transporte.Aduana.CodPaisRecep = childNodeAduana.InnerText;
                                            break;
                                        case "CodPaisDestin":
                                            objDTE.Transporte.Aduana.CodPaisDestin = childNodeAduana.InnerText;
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                }

                #endregion

                // NODO TOTALES
                #region TOTALES

                String PathTotales = "//ns:Totales";
                XmlNode Totales = xmlDoc.SelectSingleNode(PathTotales, ns);

                foreach (XmlNode childNode in Totales)
                {
                    switch (childNode.Name)
                    {
                        case "MntNeto":
                            objDTE.Totales.MntNeto = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntExe":
                            objDTE.Totales.MntExe = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntBase":
                            objDTE.Totales.MntBase = Int64.Parse(childNode.InnerText);
                            break;
                        case "MntMargenCom":
                            objDTE.Totales.MntMargenCom = Int64.Parse(childNode.InnerText);
                            break;
                        case "TasaIVA":
                            objDTE.Totales.TasaIVA = Double.Parse(childNode.InnerText.Replace(".", ","));
                            break;
                        case "IVA":
                            objDTE.Totales.IVA = Int64.Parse(childNode.InnerText);
                            break;
                        case "IVAProp":
                            objDTE.Totales.IVAProp = Int64.Parse(childNode.InnerText);
                            break;
                        case "IVATerc":
                            objDTE.Totales.IVATerc = Int64.Parse(childNode.InnerText);
                            break;
                        case "ImptoReten":
                            ImptoReten imp = new ImptoReten();
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("TipoImp")) { imp.TipoImp = child.InnerText; }
                                else if (child.Name.Equals("TasaImp")) { imp.TasaImp = Double.Parse(child.InnerText.Replace(".", ",")); }
                                else if (child.Name.Equals("MontoImp")) { imp.MontoImp = Int64.Parse(child.InnerText); }
                            }
                            objDTE.Totales.ImptoReten.Add(imp);
                            break;
                        case "IVANoRet":
                            objDTE.Totales.IVANoRet = Int64.Parse(childNode.InnerText);
                            break;
                        case "CredEC":
                            objDTE.Totales.CredEC = Int64.Parse(childNode.InnerText);
                            break;
                        case "GrntDep":
                            objDTE.Totales.GrntDep = Int64.Parse(childNode.InnerText);
                            break;
                        case "Comisiones":
                            foreach (XmlNode child in childNode.ChildNodes)
                            {
                                if (child.Name.Equals("ValComNeto")) { objDTE.Totales.ComisionesTotal.ValComNeto = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("ValComExe")) { objDTE.Totales.ComisionesTotal.ValComExe = Int64.Parse(child.InnerText); }
                                else if (child.Name.Equals("ValComIVA")) { objDTE.Totales.ComisionesTotal.ValComIVA = Int64.Parse(child.InnerText); }
                            }
                            break;
                        case "MntTotal":
                            objDTE.Totales.MntTotal = Int64.Parse(childNode.InnerText);
                            break;
                        case "MontoNF":
                            objDTE.Totales.MontoNF = Int64.Parse(childNode.InnerText);
                            break;
                        case "MontoPeriodo":
                            objDTE.Totales.MontoPeriodo = Int64.Parse(childNode.InnerText);
                            break;
                        case "SaldoAnterior":
                            objDTE.Totales.SaldoAnterior = Int64.Parse(childNode.InnerText);
                            break;
                        case "VlrPagar":
                            objDTE.Totales.VlrPagar = Int64.Parse(childNode.InnerText);
                            break;
                    }
                }

                #endregion

                // NODO OTRA MONEDA
                #region OTRAMONEDA

                String PathOtraMoneda = "//ns:OtraMoneda";
                XmlNode OtraMoneda = xmlDoc.SelectSingleNode(PathOtraMoneda, ns);

                if (OtraMoneda != null)
                {
                    foreach (XmlNode childNode in OtraMoneda)
                    {
                        switch (childNode.Name)
                        {
                            case "TpoMoneda":
                                objDTE.OtraMoneda.TpoMoneda = childNode.InnerText;
                                break;
                            case "TpoCambio":
                                objDTE.OtraMoneda.TpoCambio = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntNetoOtrMnda":
                                objDTE.OtraMoneda.MntNetoOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntExeOtrMnda":
                                objDTE.OtraMoneda.MntExeOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntFaeCarneOtrMnda":
                                objDTE.OtraMoneda.MntFaeCarneOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntMargComOtrMnda":
                                objDTE.OtraMoneda.MntMargComOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "IVAOtrMnda":
                                objDTE.OtraMoneda.IVAOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "ImpRetOtrMnda":
                                ImpRetOtrMnda imp = new ImpRetOtrMnda();
                                foreach (XmlNode child in childNode.ChildNodes)
                                {
                                    if (child.Name.Equals("TipoImpOtrMnda")) { imp.TipoImpOtrMnda = child.InnerText; }
                                    else if (child.Name.Equals("TasaImpOtrMnda")) { imp.TasaImpOtrMnda = Double.Parse(child.InnerText.Replace(".", ",")); }
                                    else if (child.Name.Equals("VlrImpOtrMnda")) { imp.VlrImpOtrMnda = Int64.Parse(child.InnerText); }
                                }
                                objDTE.OtraMoneda.ImpRetOtrMnda.Add(imp);
                                break;
                            case "IVANoRetOtrMnda":
                                objDTE.OtraMoneda.IVANoRetOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                            case "MntTotOtrMnda":
                                objDTE.OtraMoneda.MntTotOtrMnda = Double.Parse(childNode.InnerText.Replace(".", ","));
                                break;
                        }
                    }
                }

                #endregion

                // NODO DETALLE
                #region DETALLE

                String PathDetalle = "//ns:Detalle";
                XmlNodeList Detalle = xmlDoc.SelectNodes(PathDetalle, ns);

                foreach (XmlNode childNode in Detalle)
                {
                    Detalle objDetalle = new Detalle();
                    foreach (XmlNode child in childNode)
                    {
                        switch (child.Name)
                        {
                            case "NroLinDet":
                                objDetalle.NroLinDet = Int32.Parse(child.InnerText);
                                break;
                            case "CdgItem":
                                CdgItem cdg = new CdgItem();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("TpoCodigo")) { cdg.TpoCodigo = child2.InnerText; }
                                    else if (child2.Name.Equals("VlrCodigo")) { cdg.VlrCodigo = child2.InnerText; }
                                }
                                objDetalle.CdgItem.Add(cdg);
                                break;
                            case "TpoDocLiq":
                                objDetalle.TpoDocLiq = child.InnerText;
                                break;
                            case "IndExe":
                                objDetalle.IndExe = Int32.Parse(child.InnerText);
                                break;
                            case "Retenedor":
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("IndAgente")) { objDetalle.IndAgente = child2.InnerText; }
                                    else if (child2.Name.Equals("MntBaseFaena")) { objDetalle.MntBaseFaenaRet = Int64.Parse(child2.InnerText); }
                                    else if (child2.Name.Equals("MntMargComer")) { objDetalle.MntMargComer = Int64.Parse(child2.InnerText); }
                                    else if (child2.Name.Equals("PrcConsFinal")) { objDetalle.PrcConsFinal = Int64.Parse(child2.InnerText); }
                                }
                                break;
                            case "NmbItem":
                                objDetalle.NmbItem = child.InnerText;
                                break;
                            case "DscItem":
                                objDetalle.DscItem = child.InnerText;
                                break;
                            case "QtyRef":
                                objDetalle.QtyRef = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "UnmdRe":
                                objDetalle.UnmdRe = child.InnerText;
                                break;
                            case "PrcRef":
                                objDetalle.PrcRef = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "QtyItem":
                                objDetalle.QtyItem = Double.Parse(child.InnerText.Replace(".", ","));
                                break;
                            case "Subcantidad":
                                Subcantidad subCant = new Subcantidad();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("SubQty")) { subCant.SubQty = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("SubCod")) { subCant.SubCod = child2.InnerText; }
                                    else if (child2.Name.Equals("SubQty")) { subCant.SubQty = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                }
                                objDetalle.Subcantidad.Add(subCant);
                                break;
                            case "FchElabor":
                                objDetalle.FchElabor = child.InnerText;
                                break;
                            case "FchVencim":
                                objDetalle.FchVencim = child.InnerText;
                                break;
                            case "UnmdItem":
                                objDetalle.UnmdItem = child.InnerText;
                                break;
                            case "PrcItem":
                                objDetalle.PrcItem = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "OtrMnda":
                                OtrMnda otr = new OtrMnda();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("PrcOtrMon")) { otr.PrcOtrMon = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("Moneda")) { otr.Moneda = child2.InnerText; }
                                    else if (child2.Name.Equals("FctConv")) { otr.FctConv = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("DctoOtrMnda")) { otr.DctoOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("RecargoOtrMnda")) { otr.RecargoOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                    else if (child2.Name.Equals("MontoItemOtrMnda")) { otr.MontoItemOtrMnda = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                }
                                objDetalle.OtrMnda.Add(otr);
                                break;
                            case "DescuentoPct":
                                objDetalle.DescuentoPct = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "DescuentoMonto":
                                objDetalle.DescuentoMonto = Convert.ToInt64(child.InnerText);
                                break;
                            case "SubDscto":
                                SubDscto subDes = new SubDscto();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {

                                    if (child2.Name.Equals("TipoDscto")) { subDes.TipoDscto = child2.InnerText; }
                                    else if (child2.Name.Equals("ValorDscto")) { subDes.ValorDscto = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                }
                                objDetalle.SubDscto.Add(subDes);
                                break;
                            case "RecargoPct":
                                objDetalle.RecargoPct = Convert.ToDouble(child.InnerText.Replace(".", ","));
                                break;
                            case "RecargoMonto":
                                objDetalle.RecargoMonto = Convert.ToInt64(child.InnerText.Replace(".", ","));
                                break;
                            case "SubRecargo":
                                SubRecargo subRec = new SubRecargo();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("TipoRecargo")) { subRec.TipoRecargo = child2.InnerText; }
                                    else if (child2.Name.Equals("ValorRecargo")) { subRec.ValorRecargo = Double.Parse(child2.InnerText.Replace(".", ",")); }
                                }
                                objDetalle.SubRecargo.Add(subRec);
                                break;
                            case "CodImpAdic":
                                CodImpAdic cod = new CodImpAdic();
                                foreach (XmlNode child2 in child.ChildNodes)
                                {
                                    if (child2.Name.Equals("#text")) { cod.sCodImpAdic = child2.InnerText; }
                                }
                                objDetalle.CodImpAdic.Add(cod);
                                break;
                            case "MontoItem":
                                objDetalle.MontoItem = Convert.ToInt64(child.InnerText.Replace(".", ","));
                                break;
                        }
                    }
                    objDTE.Detalle.Add(objDetalle);
                }

                #endregion

                // NODO SUBTOTALES INFORMATIVOS
                #region SUBTOTALES INFORMATIVOS

                String PathSubtotales = "//ns:SubTotInfo";
                XmlNodeList SubTotInfo = xmlDoc.SelectNodes(PathSubtotales, ns);

                if (SubTotInfo != null)
                {
                    foreach (XmlNode childNode in SubTotInfo)
                    {
                        SubTotInfo objSubTotInfo = new SubTotInfo();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroSTI":
                                    objSubTotInfo.NroSTI = Int32.Parse(child.InnerText);
                                    break;
                                case "GlosaSTI":
                                    objSubTotInfo.GlosaSTI = child.InnerText;
                                    break;
                                case "OrdenSTI":
                                    objSubTotInfo.OrdenSTI = Int32.Parse(child.InnerText);
                                    break;
                                case "SubTotNetoSTI":
                                    objSubTotInfo.SubTotNetoSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotIVASTI":
                                    objSubTotInfo.SubTotIVASTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotAdicSTI":
                                    objSubTotInfo.SubTotAdicSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "SubTotExeSTI":
                                    objSubTotInfo.SubTotExeSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "ValSubtotSTI":
                                    objSubTotInfo.ValSubtotSTI = Double.Parse(child.InnerText.Replace(".", ","));
                                    break;
                                case "LineasDeta":
                                    LineasDeta linDet = new LineasDeta();
                                    foreach (XmlNode child2 in child.ChildNodes)
                                    {
                                        if (child.Name.Equals("LineasDeta")) { linDet.iLineasDeta = Int32.Parse(child2.InnerText); }
                                    }
                                    objSubTotInfo.LineasDeta.Add(linDet);
                                    break;
                            }
                        }
                        objDTE.SubTotInfo.Add(objSubTotInfo);
                    }
                }

                #endregion

                // NODO DESCUENTOS Y/O RECARGOS
                #region DESCUENTOS y/o RECARGOS

                String PathDscRcgGlobal = "//ns:DscRcgGlobal";
                XmlNodeList DscRcgGlobal = xmlDoc.SelectNodes(PathDscRcgGlobal, ns);

                if (DscRcgGlobal != null)
                {
                    foreach (XmlNode childNode in DscRcgGlobal)
                    {
                        DscRcgGlobal objDscRcgGlobal = new DscRcgGlobal();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinDR":
                                    objDscRcgGlobal.NroLinDR = Int32.Parse(child.InnerText);
                                    break;
                                case "TpoMov":
                                    objDscRcgGlobal.TpoMov = child.InnerText;
                                    break;
                                case "GlosaDR":
                                    objDscRcgGlobal.GlosaDR = child.InnerText;
                                    break;
                                case "TpoValor":
                                    objDscRcgGlobal.TpoValor = child.InnerText;
                                    break;
                                case "ValorDR":
                                    objDscRcgGlobal.ValorDR = Double.Parse(child.InnerText);
                                    break;
                                case "ValorDROtrMnda":
                                    objDscRcgGlobal.ValorDROtrMnda = Double.Parse(child.InnerText);
                                    break;
                                case "IndExeDR":
                                    objDscRcgGlobal.IndExeDR = Int32.Parse(child.InnerText);
                                    break;
                            }
                        }
                        objDTE.DscRcgGlobal.Add(objDscRcgGlobal);
                    }
                }

                #endregion

                // NODO REFERENCIAS
                #region REFERENCIAS

                String PathReferencia = "//ns:Referencia";
                XmlNodeList Referencia = xmlDoc.SelectNodes(PathReferencia, ns);

                if (Referencia != null)
                {
                    foreach (XmlNode childNode in Referencia)
                    {
                        Referencia objReferencia = new Referencia();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinRef":
                                    objReferencia.NroLinRef = Int32.Parse(child.InnerText);
                                    break;
                                case "TpoDocRef":
                                    objReferencia.TpoDocRef = child.InnerText;
                                    break;
                                case "IndGlobal":
                                    objReferencia.IndGlobal = Int32.Parse(child.InnerText);
                                    break;
                                case "FolioRef":
                                    objReferencia.FolioRef = child.InnerText;
                                    break;
                                case "RUTOtr":
                                    objReferencia.RUTOtr = child.InnerText;
                                    break;
                                case "FchRef":
                                    objReferencia.FchRef = child.InnerText;
                                    break;
                                case "CodRef":
                                    objReferencia.CodRef = Int32.Parse(child.InnerText);
                                    break;
                                case "RazonRef":
                                    objReferencia.RazonRef = child.InnerText;
                                    break;
                            }
                        }
                        objDTE.Referencia.Add(objReferencia);
                    }
                }

                #endregion

                // NODO COMISIONES
                #region COMISIONES

                String PathComisiones = "//ns:Comisiones";
                XmlNodeList Comisiones = xmlDoc.SelectNodes(PathComisiones, ns);

                if (Comisiones != null)
                {
                    foreach (XmlNode childNode in Comisiones)
                    {
                        Comisiones objComisiones = new Comisiones();
                        foreach (XmlNode child in childNode)
                        {
                            switch (child.Name)
                            {
                                case "NroLinCom":
                                    objComisiones.NroLinCom = Int32.Parse(child.InnerText);
                                    break;
                                case "TipoMovim":
                                    objComisiones.TipoMovim = child.InnerText;
                                    break;
                                case "Glosa":
                                    objComisiones.Glosa = child.InnerText;
                                    break;
                                case "TasaComision":
                                    objComisiones.TasaComision = Double.Parse(child.InnerText);
                                    break;
                                case "ValComNeto":
                                    objComisiones.ValComNeto = Int64.Parse(child.InnerText);
                                    break;
                                case "ValComExe":
                                    objComisiones.ValComExe = Int64.Parse(child.InnerText);
                                    break;
                                case "ValComIVA":
                                    objComisiones.ValComIVA = Int64.Parse(child.InnerText);
                                    break;
                            }
                        }
                        objDTE.Comisiones.Add(objComisiones);
                    }
                }

                #endregion

                resultDTE.Success = true;
                resultDTE.DTE = objDTE;
                resultDTE.XMLString = decodeStringXml;
                return resultDTE;
            }
            catch (Exception ex)
            {

                resultDTE.Success = false;
                resultDTE.Mensaje = ex.Message;
                return resultDTE;
            }
        }

        private SearchResultXML GetSearchResultXML(string searchResultString)
        {
            SearchResultXML searchResultXML = new SearchResultXML();
            //string StringToXml = Base64Decode(searchResultString);
            try
            {
                XmlDocument xmlSearchResul = new XmlDocument();                
                xmlSearchResul.LoadXml(searchResultString);
                //string FD = xmlSearchResul.ChildNodes.Item(0);// new System.Linq.SystemCore_EnumerableDebugView(xmlSearchResul.ChildNodes.Item(0)).Items[0];
                //XmlNamespaceManager ns = new XmlNamespaceManager(xmlSearchResul.NameTable);
                //String PathSearchResult = "SearchResult";
                //XmlNode SearchResul = xmlSearchResul.SelectSingleNode(PathSearchResult, ns);
                XmlNodeList nodeList = xmlSearchResul.DocumentElement.ChildNodes;
                foreach (XmlNode xmlNode in nodeList)
                {
                    switch (xmlNode.Name)
                    {
                        case "Data":
                            string DataDecode = Base64Decode(xmlNode.InnerText);
                            XmlDocument DataXML = new XmlDocument();
                            DataXML.LoadXml(DataDecode);
                            XmlNodeList DataNodeList = DataXML.DocumentElement.ChildNodes;
                            Data data = new Data();
                            //List<DocumentXML> documentXMLs = new List<DocumentXML>();
                            DocumentXML document = new DocumentXML();
                            foreach (XmlElement xmlElement in DataXML.DocumentElement)
                            {
                                document = new DocumentXML();
                                DataNodeList = xmlElement.ChildNodes;
                                foreach (XmlNode NodeData in DataNodeList)
                                {
                                    switch (NodeData.Name)
                                    {
                                        case "RecipientRUT":
                                            document.RecipientRUT = NodeData.InnerText;
                                            break;
                                        case "IssuerName":
                                            document.IssuerName = NodeData.InnerText;
                                            break;
                                        case "Created":
                                            document.Created = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "IssuerRUT":
                                            document.IssuerRUT = NodeData.InnerText;
                                            break;
                                        case "ExternalID":
                                            document.ExternalID = NodeData.InnerText;
                                            break;
                                        case "Replaced":
                                            document.Replaced = Boolean.Parse(NodeData.InnerText);
                                            break;
                                        case "GetUniqueBusinessId":
                                            document.GetUniqueBusinessId = NodeData.InnerText;
                                            break;
                                        case "DocumentReferences":
                                            document.DocumentReferences = NodeData.InnerText;
                                            break;
                                        case "Statuses":
                                            document.Statuses = NodeData.InnerText;
                                            break;
                                        case "Version":
                                            document.Version = NodeData.InnerText;
                                            break;
                                        case "TotalAmount":
                                            document.TotalAmount = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "Deleted":
                                            document.Deleted = Boolean.Parse(NodeData.InnerText);
                                            break;
                                        case "EmissionDate":
                                            document.EmissionDate = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "MailReferences":
                                            document.MailReferences = NodeData.InnerText;
                                            break;
                                        case "DTEType":
                                            document.DTEType = NodeData.InnerText;
                                            break;
                                        case "Iva":
                                            document.Iva = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "Folio":
                                            document.Folio = NodeData.InnerText;
                                            break;
                                        case "RecipientName":
                                            document.RecipientName = NodeData.InnerText;
                                            break;
                                        case "Modified":
                                            document.Modified = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "DocumentID":
                                            document.DocumentID = NodeData.InnerText;
                                            break;
                                        case "Id":
                                            document.Id = NodeData.InnerText;
                                            break;
                                        case "FchRespComercial":
                                            document.FchRespComercial = NodeData.InnerText;
                                            break;
                                        case "NetoAmount":
                                            document.NetoAmount = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "CanalRecepcion":
                                            document.CanalRecepcion = NodeData.InnerText;
                                            break;
                                        case "Anulado":
                                            document.Anulado = NodeData.InnerText;
                                            break;
                                        case "Eliminado":
                                            document.Eliminado = NodeData.InnerText;
                                            break;
                                        case "FchRecepSII":
                                            document.FchRecepSII = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "Firma":
                                            document.Firma = NodeData.InnerText;
                                            break;
                                        case "Desanulado":
                                            document.Desanulado = NodeData.InnerText;
                                            break;
                                        case "Intercambio":
                                            document.Intercambio = NodeData.InnerText;
                                            break;
                                        case "AutorizadoSII":
                                            document.AutorizadoSII = NodeData.InnerText;
                                            break;
                                        case "Recibido":
                                            document.Recibido = NodeData.InnerText;
                                            break;
                                        case "Distribuido":
                                            document.Distribuido = NodeData.InnerText;
                                            break;
                                        case "CmnaRecep":
                                            document.CmnaRecep = NodeData.InnerText;
                                            break;
                                        case "TieneArchivo":
                                            document.TieneArchivo = Boolean.Parse(NodeData.InnerText);
                                            break;
                                        case "AnuladoContable":
                                            document.AnuladoContable = NodeData.InnerText;
                                            break;
                                        case "RUTRecep":
                                            document.RUTRecep = NodeData.InnerText;
                                            break;
                                        case "Grupo":
                                            document.Grupo = NodeData.InnerText;
                                            break;
                                        case "Elaboracion":
                                            document.Elaboracion = NodeData.InnerText;
                                            break;
                                        case "MntNeto":
                                            document.MntNeto = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "ID":
                                            document.ID = NodeData.InnerText;
                                            break;
                                        case "CiudadOrigen":
                                            document.CiudadOrigen = NodeData.InnerText;
                                            break;
                                        case "MntTotal":
                                            document.MntTotal = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "Estructura":
                                            document.Estructura = NodeData.InnerText;
                                            break;
                                        case "FmaPago":
                                            document.FmaPago = int.Parse(NodeData.InnerText);
                                            break;
                                        case "FchEmis":
                                            document.FchEmis = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "ErrorPrint":
                                            document.ErrorPrint = NodeData.InnerText;
                                            break;
                                        case "CmnaOrigen":
                                            document.CmnaOrigen = NodeData.InnerText;
                                            break;
                                        case "CdgSIISucur":
                                            document.CdgSIISucur = NodeData.InnerText;
                                            break;
                                        case "Contacto":
                                            document.Contacto = NodeData.InnerText;
                                            break;
                                        case "CiudadRecep":
                                            document.CiudadRecep = NodeData.InnerText;
                                            break;
                                        case "RUTEmisor":
                                            document.RUTEmisor = NodeData.InnerText;
                                            break;
                                        case "CEN":
                                            document.CEN = NodeData.InnerText;
                                            break;
                                        case "Aprobado":
                                            document.Aprobado = NodeData.InnerText;
                                            break;
                                        case "Cesion":
                                            document.Cesion = NodeData.InnerText;
                                            break;
                                        case "TimeStamp":
                                            document.TimeStamp = DateTime.Parse(NodeData.InnerText);
                                            break;
                                        case "TipoDocumento":
                                            document.TipoDocumento = NodeData.InnerText;
                                            break;
                                        case "AprobadoSII":
                                            document.AprobadoSII = NodeData.InnerText;
                                            break;
                                        case "IVA":
                                            document.IVA = Double.Parse(NodeData.InnerText);
                                            break;
                                        case "RznSoc":
                                            document.RznSoc = NodeData.InnerText;
                                            break;
                                        case "TipoDTE":
                                            document.TipoDTE = NodeData.InnerText;
                                            break;
                                        case "NmbItem":
                                            document.NmbItem = NodeData.InnerText;
                                            break;
                                        case "Conciliado":
                                            document.Conciliado = NodeData.InnerText;
                                            break;
                                        case "Procesado":
                                            document.Procesado = NodeData.InnerText;
                                            break;
                                        case "RznSocRecep":
                                            document.RznSocRecep = NodeData.InnerText;
                                            break;
                                        case "DscItem":
                                            document.DscItem = NodeData.InnerText;
                                            break;
                                        case "DownloadCustomerDocumentUrl":
                                            document.DownloadCustomerDocumentUrl = NodeData.InnerText;
                                            break;
                                        case "ClaimAction":
                                            document.ClaimAction = NodeData.InnerText;
                                            break;
                                    }
                                }
                                data.ltDocuments.Add(document);
                            }
                            data.stringData = DataDecode;
                            
                            searchResultXML.Data = data;
                            break;
                        case "Description":
                            searchResultXML.Description = xmlNode.InnerText;
                            break;
                        case "Result":
                            searchResultXML.Result = xmlNode.InnerText;
                            break;
                        case "StackTrace":
                            searchResultXML.StackTrace = xmlNode.InnerText;
                            break;
                        case "TotalDocuments":
                            searchResultXML.TotalDocuments = Int32.Parse(xmlNode.InnerText);
                            break;
                    }
                }
            }
#pragma warning disable CS0168 // La variable 'ex' se ha declarado pero nunca se usa
            catch (Exception ex)
#pragma warning restore CS0168 // La variable 'ex' se ha declarado pero nunca se usa
            {

            }
            return searchResultXML;
        }

        public string Base64Encode(string plainText)
        {
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }

    public class DTECompany
    {
        public  string ENVIRONMENT { get; set; }
        public  string RUTREC { get; set; }
        public  bool SENTACD { get; set; }
        public  bool SENTRZD { get; set; }
        public  string URLDTELIST { get; set; }
        public  string URLDTE { get; set; }
        public  string URLACD { get; set; }
        public  string URLRZD { get; set; }
        public  string USER { get; set; }
        public  string KEY { get; set; }
        public string LKPRCOR { get; set; }
        public string LKPRCDN { get; set; }
    }
}
