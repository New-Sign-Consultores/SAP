using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common.CRUD
{
    public class CRUD_ASRDTE
    {
        public CRUD_ASRDTE()
        {

        }

        public void AddASRDTE(ASRDTE aSRDTE)
        {
            string Code = null;
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.CompanyService oCompService;
                oCompService = Conex.oCompany.GetCompanyService();
                oGeneralService = oCompService.GetGeneralService("ASRDTE");
               
                Code = ExistASRDTE(aSRDTE.RutEmisor, aSRDTE.TipoDTE, aSRDTE.Folio);
                if (string.IsNullOrEmpty(Code))
                {
                    SAPbobsCOM.GeneralService oDocGeneralService = null;
                    SAPbobsCOM.GeneralData oDocGeneralData = null;
                    try
                    {
                        try
                        {
                            #region fields
                            Conex.oCompany.StartTransaction();
                            //oCompService = Conexiones.ConexSAP.oCompany.GetCompanyService();
                            oDocGeneralService = oCompService.GetGeneralService("ASRDTE");
                            oDocGeneralData = (SAPbobsCOM.GeneralData)oDocGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                            oDocGeneralData.SetProperty("U_DocumentID", aSRDTE.DocumentID);
                            oDocGeneralData.SetProperty("U_RutEmisor", aSRDTE.RutEmisor);
                            oDocGeneralData.SetProperty("U_RznSoc", aSRDTE.RznSoc);
                            oDocGeneralData.SetProperty("U_TipoDTE", aSRDTE.TipoDTE);
                            oDocGeneralData.SetProperty("U_ExiEmisor", (aSRDTE.ExiEmisor ? "Y" : "N"));
                            oDocGeneralData.SetProperty("U_Folio", aSRDTE.Folio);
                            oDocGeneralData.SetProperty("U_FchEmis", aSRDTE.FchEmis);
                            oDocGeneralData.SetProperty("U_FchVenc", aSRDTE.FchVenc);
                            oDocGeneralData.SetProperty("U_FmaPago", aSRDTE.FmaPago);
                            oDocGeneralData.SetProperty("U_MntNeto", aSRDTE.MntNeto);
                            oDocGeneralData.SetProperty("U_MntExe", aSRDTE.MntExe);
                            oDocGeneralData.SetProperty("U_TasaIVA", aSRDTE.TasaIVA);
                            oDocGeneralData.SetProperty("U_IVA", aSRDTE.IVA);
                            oDocGeneralData.SetProperty("U_MntTotal", aSRDTE.MntTotal);
                            oDocGeneralData.SetProperty("U_FolioRefOC", aSRDTE.FolioRefOC);
                            oDocGeneralData.SetProperty("U_FolioSAPOC", aSRDTE.FolioSAPOC);
                            oDocGeneralData.SetProperty("U_FolioRefEM", aSRDTE.FolioRefEM);
                            oDocGeneralData.SetProperty("U_FolioSAPEM", aSRDTE.FolioSAPEM);
                            oDocGeneralData.SetProperty("U_FolioRefFA", aSRDTE.FolioRefFA);
                            oDocGeneralData.SetProperty("U_FolioSAPFA", aSRDTE.FolioSAPFA);
                            oDocGeneralData.SetProperty("U_DocEntryS", aSRDTE.DocEntryS);
                            oDocGeneralData.SetProperty("U_ObjType", aSRDTE.ObjType);
                            oDocGeneralData.SetProperty("U_XML", aSRDTE.XML);
                            oDocGeneralData.SetProperty("U_PDF64", aSRDTE.PDF64);

                            #endregion fields

                            oDocGeneralService.Add(oDocGeneralData);
                            if (Conex.oCompany.InTransaction)
                            {
                                Conex.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompService);
                            oCompService = null;
                            GC.Collect();
                        }
                        catch (Exception ex)
                        {
                            if (!ex.Message.Contains("Esta entrada ya existe en las tablas siguientes"))
                                //    Comunes.LogArchivo.EscribeLog(ex);
                                if (Conex.oCompany.InTransaction)
                            {
                                Conex.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                        }

                    }
                    finally
                    {
                        GC.Collect();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocGeneralService);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDocGeneralData);
                        oDocGeneralService = null;
                        oCompService = null;
                        oDocGeneralData = null;
                    }
                }

            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("CRDTE:AddASRDTE : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public bool UpdateASRDTE(string Code, string DocEntryS, string ObjType)
        {
            bool flag = false;
            SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                string query = @" UPDATE T0 
                                SET T0.""U_DocEntryS"" = '" + DocEntryS + @"'
                                   ,T0.""U_ObjType"" = '" + ObjType + @"'
                                FROM ""@ASRDTE"" T0 WHERE T0.""Code"" = '" + Code + @"' ";
                recordset.DoQuery(query);
                flag = true;
            }
            finally
            {
               
            }
            return flag;
        }

        public string ExistASRDTE(string RUT, string Tipo, string Folio)
       {
            string Code = null;
            //bool Existe = false;
            try
            {                
               
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Query = @" SELECT ""Code"" FROM ""@ASRDTE"" T0 WHERE  T0.""U_RutEmisor"" = '" + RUT + @"' AND T0.""U_TipoDTE"" = '" + Tipo + @"' AND T0.""U_Folio"" = '" + Folio + @"' ";
                recordset.DoQuery(Query);
                if(recordset.RecordCount > 0)
                {
                    Code = Convert.ToString(recordset.Fields.Item("Code").Value);
                }
            }
            catch(Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("CRDTE:ExistASRDTE : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return Code;
       }

        public ResultSt ExistSAP(string RUT, string Tipo, string Folio)
        {
            ResultSt resultSt = new ResultSt();
            try
            {
                string Table = null;
                switch (Tipo)
                {
                    case "33":
                        Table = "OPCH";
                        break;
                    case "61":
                        Table = "ORPC";
                        break;
                }
                SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string Query = @" SELECT T0.""DocEntry"", T0.""ObjType"" FROM " + Table + @" T0 WHERE  T0.""LicTradNum"" = '" + RUT + @"' AND T0.""Indicator"" = " + Tipo + @" AND T0.""FolioNum"" = " + Folio + @" ";
                recordset.DoQuery(Query);
                if (recordset.RecordCount > 0)
                {
                    resultSt.Codigo= Convert.ToString(recordset.Fields.Item("DocEntry").Value); 
                    resultSt.Valor = Convert.ToString(recordset.Fields.Item("ObjType").Value);
                    resultSt.Estado = true;

                }
            }
            catch (Exception ex)
            {
                resultSt.Estado = false;
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("CRDTE:ExistSAP : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return resultSt;
        }
    }
}
