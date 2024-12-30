using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common
{
    public class Common
    {
        CultureInfo cultureInfo = new CultureInfo("es-CL");
        public int greenForeColor = 3329330;
        public int redBackColor = 255;
        public Common()
        {

        }

        public ResultRefSAP GetRefOC(string foliosOC, string RutEmisor, string LkPrcOr)
        {
            ResultRefSAP resultRefSAP = new ResultRefSAP();
            try
            {
                string query = null;
                query = @" SELECT SUM(T0.DocTotal) AS DocTotal, 
                                Convert(VARCHAR(100),STUFF((
                                    SELECT ','+ CONVERT(VARCHAR(30),T0.DocNum)
                                    FROM OPOR T0 WHERE CONVERT(VARCHAR(100), T0." + LkPrcOr + @") IN(" + foliosOC + @")  AND T0.LicTradNum = '" + RutEmisor + @"'
                                    AND T0.CANCELED <> 'Y'
                                        FOR XML PATH('')

                                    ),1,1, '')) AS DocNum
                                    FROM OPOR T0
                                    WHERE CONVERT(VARCHAR(100), T0." + LkPrcOr + @") IN(" + foliosOC + @")  AND T0.LicTradNum = '" + RutEmisor + @"'
                                    AND T0.DocStatus != 'C'
                                    AND T0.CANCELED <> 'Y'
                                    GROUP BY T0.DocType ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    resultRefSAP.DocNum = Convert.ToString( recordset.Fields.Item("DocNum").Value);
                    resultRefSAP.DocTotal = Convert.ToDouble(recordset.Fields.Item("DocTotal").Value);
                    resultRefSAP.existe = true;
                }
                else resultRefSAP.existe = false;

            }
            catch (Exception ex)
            {
                resultRefSAP.existe = false;
                throw ex;

            }
            return resultRefSAP;                     
        }

        public ResultRefSAP GetRefEM(string foliosEM, string RutEmisor, string LkPrcDn)
        {
            ResultRefSAP resultRefSAP = new ResultRefSAP();
            try
            {
                string query = null;
                query = @" SELECT SUM(T0.DocTotal) AS DocTotal, 
                                Convert(VARCHAR(100),STUFF((
                                    SELECT ','+ CONVERT(VARCHAR(30),T0.DocNum)
                                    FROM OPDN T0 WHERE CONVERT(VARCHAR(100), T0." + LkPrcDn + @") IN(" + foliosEM + @")  AND T0.LicTradNum = '" + RutEmisor + @"'
                                    AND T0.CANCELED <> 'Y'
                                        FOR XML PATH('')

                                    ),1,1, '')) AS DocNum
                                    FROM OPDN T0
                                    WHERE CONVERT(VARCHAR(100), T0." + LkPrcDn + @") IN(" + foliosEM + @")  AND T0.LicTradNum = '" + RutEmisor + @"'
                                    AND T0.DocStatus != 'C'
                                    AND T0.CANCELED <> 'Y'
                                    GROUP BY T0.DocType ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    resultRefSAP.DocNum = Convert.ToString(recordset.Fields.Item("DocNum").Value);
                    resultRefSAP.DocTotal = Convert.ToDouble(recordset.Fields.Item("DocTotal").Value);
                    resultRefSAP.existe = true;
                }
                else resultRefSAP.existe = false;

            }
            catch (Exception ex)
            {
                resultRefSAP.existe = false;
                throw ex;

            }
            return resultRefSAP;
        }

        public ResultRefSAP GetRefFA(string foliosFA, string RutEmisor)
        {
            ResultRefSAP resultRefSAP = new ResultRefSAP();
            try
            {
                string query = null;
                query = @" SELECT SUM(T0.DocTotal) AS DocTotal, 
                                Convert(VARCHAR(100),STUFF((
                                    SELECT ','+ CONVERT(VARCHAR(30),T0.DocNum)
                                    FROM OPCH T0 WHERE CONVERT(VARCHAR(100), T0.FolioNum) IN(" + foliosFA + @")  AND T0.LicTradNum = '" + RutEmisor + @"'
                                    AND T0.CANCELED <> 'Y'
                                        FOR XML PATH('')

                                    ),1,1, '')) AS DocNum
                                    FROM OPCH T0
                                    WHERE CONVERT(VARCHAR(100), T0.FolioNum) IN(" + foliosFA + @")  AND T0.LicTradNum = '" + RutEmisor + @"'                                
                                    AND T0.CANCELED <> 'Y'
                                    GROUP BY T0.DocType ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    resultRefSAP.DocNum = Convert.ToString(recordset.Fields.Item("DocNum").Value);
                    resultRefSAP.DocTotal = Convert.ToDouble(recordset.Fields.Item("DocTotal").Value);
                    resultRefSAP.existe = true;
                }
                else resultRefSAP.existe = false;

            }
            catch (Exception ex)
            {
                resultRefSAP.existe = false;
                throw ex;

            }
            return resultRefSAP;
        }

        public List<DocDetSAP> GetDtRefOC(string foliosOC, string RutEmisor)
        {
            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
            DocDetSAP docDetSAP = new DocDetSAP();
            try
            {
                string query = null;
                query = @" SELECT T1.""DocEntry""
	                            ,T1.""LineNum""
	                            ,T0.""ObjType""
	                            ,T1.""Currency""
	                            ,T1.""Rate""
	                            ,T0.""DocCur""
	                            ,T0.""DocRate""
                                ,T0.""DocType""
                            FROM OPOR T0
                            INNER JOIN POR1 T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            WHERE T0.""DocNum"" IN (" + foliosOC + @")
	                            AND T0.""LicTradNum"" = '" + RutEmisor + @"' 
                                AND T0.""DocStatus"" != 'C' 
                                AND T0.""CANCELED"" != 'Y'
                                AND T1.""LineStatus"" = 'O' ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    for (int i = 0; i < recordset.RecordCount; i++)
                    {
                        docDetSAP = new DocDetSAP();
                        docDetSAP.DocEntry = Convert.ToInt32(recordset.Fields.Item("DocEntry").Value);
                        docDetSAP.LineNum = Convert.ToInt32(recordset.Fields.Item("LineNum").Value);
                        docDetSAP.ObjType = Convert.ToString(recordset.Fields.Item("ObjType").Value);
                        docDetSAP.Currency = Convert.ToString(recordset.Fields.Item("Currency").Value);
                        docDetSAP.Rate = Convert.ToDouble(recordset.Fields.Item("Rate").Value);
                        docDetSAP.DocCur = Convert.ToString(recordset.Fields.Item("DocCur").Value);
                        docDetSAP.DocRate = Convert.ToDouble(recordset.Fields.Item("DocRate").Value);
                        docDetSAP.DocType = Convert.ToString(recordset.Fields.Item("DocType").Value);
                        recordset.MoveNext();
                        docDetSAPs.Add(docDetSAP);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;

            }
            return docDetSAPs;
        }

        public List<DocDetSAP> GetDtRefEM(string foliosEM, string RutEmisor)
        {
            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
            DocDetSAP docDetSAP = new DocDetSAP();
            try
            {
                string query = null;
                query = @" SELECT T1.""DocEntry""
	                            ,T1.""LineNum""
	                            ,T0.""ObjType""
	                            ,T1.""Currency""
	                            ,T1.""Rate""
	                            ,T0.""DocCur""
	                            ,T0.""DocRate""
                                ,T0.""DocType""
                            FROM OPDN T0
                            INNER JOIN PDN1 T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            WHERE T0.""DocNum"" IN (" + foliosEM + @")
	                            AND T0.""LicTradNum"" = '" + RutEmisor + @"'
                                AND T0.""DocStatus"" != 'C' 
                                AND T0.""CANCELED"" != 'Y' 
                                AND T1.""LineStatus"" = 'O' ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    for (int i = 0; i < recordset.RecordCount; i++)
                    {
                        docDetSAP.DocEntry = Convert.ToInt32(recordset.Fields.Item("DocEntry").Value);
                        docDetSAP.LineNum = Convert.ToInt32(recordset.Fields.Item("LineNum").Value);
                        docDetSAP.ObjType = Convert.ToString(recordset.Fields.Item("ObjType").Value);
                        docDetSAP.Currency = Convert.ToString(recordset.Fields.Item("Currency").Value);
                        docDetSAP.Rate = Convert.ToDouble(recordset.Fields.Item("Rate").Value);
                        docDetSAP.DocCur = Convert.ToString(recordset.Fields.Item("DocCur").Value);
                        docDetSAP.DocRate = Convert.ToDouble(recordset.Fields.Item("DocRate").Value);
                        docDetSAP.DocType = Convert.ToString(recordset.Fields.Item("DocType").Value);
                        recordset.MoveNext();
                        docDetSAPs.Add(docDetSAP);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;

            }
            return docDetSAPs;
        }

        public List<DocDetSAP> GetDtRefFA(string foliosFA, string RutEmisor)
        {
            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
            DocDetSAP docDetSAP = new DocDetSAP();
            try
            {
                string query = null;
                query = @" SELECT T1.""DocEntry""
	                            ,T1.""LineNum""
	                            ,T0.""ObjType""
	                            ,T1.""Currency""
	                            ,T1.""Rate""
	                            ,T0.""DocCur""
	                            ,T0.""DocRate""
                            FROM OPCH T0
                            INNER JOIN PCH1 T1 ON T0.""DocEntry"" = T1.""DocEntry""
                            WHERE T0.""DocNum"" IN (" + foliosFA + @")
	                            AND T0.""LicTradNum"" = '" + RutEmisor + @"'
                                AND T0.""CANCELED"" != 'Y' ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    for (int i = 0; i < recordset.RecordCount; i++)
                    {
                        docDetSAP.DocEntry = Convert.ToInt32(recordset.Fields.Item("DocEntry").Value);
                        docDetSAP.LineNum = Convert.ToInt32(recordset.Fields.Item("LineNum").Value);
                        docDetSAP.ObjType = Convert.ToString(recordset.Fields.Item("ObjType").Value);
                        docDetSAP.Currency = Convert.ToString(recordset.Fields.Item("Currency").Value);
                        docDetSAP.Rate = Convert.ToDouble(recordset.Fields.Item("Rate").Value);
                        docDetSAP.DocCur = Convert.ToString(recordset.Fields.Item("DocCur").Value);
                        docDetSAP.DocRate = Convert.ToDouble(recordset.Fields.Item("DocRate").Value);
                        recordset.MoveNext();
                        docDetSAPs.Add(docDetSAP);
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;

            }
            return docDetSAPs;
        }

        public int DocInTable(string FchDesde, string FchHasta, string TipoDoc)
        {
            int cantDoc = 0;
            try
            {
                string Query = Properties.Settings.Default.ListaDTESQL;

                #region Rango de fecha
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
                Query = Query.Replace("FchDesd", "'" + FchDesd.ToString("yyyyMMdd") + "'");
                Query = Query.Replace("FchHast", "'" + FchHast.ToString("yyyyMMdd") + "'");
                #endregion Rango de fecha

                #region Tipo documento
                Query = Query.Replace("Documentos", "'" + TipoDoc + "'");
                #endregion Tipo documento

                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(Query);
                if (!recordset.EoF)
                {
                    cantDoc = recordset.RecordCount;
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
            return cantDoc;
        }

        public BPSAP GetBPSAP(string RutEmisor)
        {
            BPSAP bPSAP = new BPSAP();
            try
            {
                string query = null;
                query = @" SELECT 
                             T0.""CardCode""
                            ,T0.""CardName""
                            ,T0.""LicTradNum"" 
                            FROM OCRD T0 
                            WHERE T0.[LicTradNum]  = '" + RutEmisor + @"' 
                            AND T0.[CardType]  ='S' ";
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));
                recordset.DoQuery(query);
                if (!recordset.EoF)
                {
                    bPSAP.CardCode = Convert.ToString(recordset.Fields.Item("CardCode").Value);
                    bPSAP.CardName = Convert.ToString(recordset.Fields.Item("CardName").Value);
                    bPSAP.LicTradNum = Convert.ToString(recordset.Fields.Item("LicTradNum").Value);
                }
                else
                    bPSAP.existe = false;

            }
            catch (Exception ex)
            {
                bPSAP.existe = false;
                throw ex;

            }
            return bPSAP;
        }
    }
    public class ResultRefSAP
    {
        public bool existe { get; set; }
        public string DocNum { get; set; }
        public double DocTotal { get; set; }
        public double DocTotalFC { get; set; }
        public string DocCur { get; set; }

        public ResultRefSAP()
        {

        }
    }

    public class DocDetSAP
    {
        public int DocEntry { get; set; }
        public int LineNum { get; set; }
        public string ObjType { get; set; }
        public string Currency { get; set; }
        public double Rate { get; set; }
        public string DocCur { get; set; }
        public double DocRate { get; set; }
        public string DocType { get; set; }
        public DocDetSAP()
        {

        }
    }

    public class BPSAP
    {
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public string LicTradNum { get; set; }
        public bool existe { get; set; }       
    }
}
