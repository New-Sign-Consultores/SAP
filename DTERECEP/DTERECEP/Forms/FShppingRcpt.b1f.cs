using DTERECEP.Common;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DTERECEP.Forms
{
    [FormAttribute("DTERECEP.Forms.FShppingRcpt", "Forms/FShppingRcpt.b1f")]
    class FShppingRcpt : UserFormBase
    {
        CultureInfo cultureInfo = new CultureInfo("es-CL");
        Common.Common common = new Common.Common();
        public FShppingRcpt()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lFchDesd").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("TFchDesd").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Consultar").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("MDTE").Specific));
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.LinkPressedAfter += new SAPbouiCOM._IMatrixEvents_LinkPressedAfterEventHandler(this.Matrix0_LinkPressedAfter);
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lFchHast").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("TFchHast").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lTipoDoc").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("CTipoDoc").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lTInt").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("TotalInt").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lTotalNInt").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("TotalNInt").Specific));
            this.ButtonCombo0 = ((SAPbouiCOM.ButtonCombo)(this.GetItem("obProcess").Specific));
            this.ButtonCombo0.ComboSelectAfter += new SAPbouiCOM._IButtonComboEvents_ComboSelectAfterEventHandler(this.ButtonCombo0_ComboSelectAfter);
            this.ButtonCombo0.ClickAfter += new SAPbouiCOM._IButtonComboEvents_ClickAfterEventHandler(this.ButtonCombo0_ClickAfter);
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("ReserInv").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void FillOpcion()
        {
            try
            {
                SAPbouiCOM.ButtonCombo buttonCombo = (SAPbouiCOM.ButtonCombo)this.UIAPIRawForm.Items.Item("obProcess").Specific;
                buttonCombo.ValidValues.Add("Con_OC", "Integrar con orden de compra");
                buttonCombo.ValidValues.Add("Con_EM", "Integrar con entrada de mercancia");
                buttonCombo.ValidValues.Add("Como_Ser", "Integrar como servicio");
                buttonCombo.ValidValues.Add("Fill_BP", "Llenar proveedor");
                buttonCombo.ValidValues.Add("Int_NC", "Integrar Nota de crédito");
                buttonCombo.ValidValues.Add("Int_ND", "Integrar Nota de débito");
                buttonCombo.Select("Con_OC", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //ComboBox1.Item.Visible = true;
                buttonCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            }
            catch
            {

            }
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            Matrix0.CommonSetting.FixedColumnsCount = 3;
            FillOpcion();
            ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            //ComboBox1.Item.Visible = false;

            ComboBox0.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
        } 

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Matrix Matrix0;

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string FchDesde = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TFchDesd").Specific).String;
            string FchHasta = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TFchHast").Specific).String;
            string TipoDoc = ((SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("CTipoDoc").Specific).Value;
            try
            {
                Common.DTENewSign dTENewSign = new Common.DTENewSign();                
                if (string.IsNullOrEmpty(TipoDoc))
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar el tipo de documento.!", 1,"OK");
                }
                else
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Buscando documentos. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    //dTENewSign.GETListDTE(FchDesde, FchHasta, TipoDoc);
                    FillLisDTE(FchDesde, FchHasta, TipoDoc);
                }
            }
            catch(Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("GETListDTE: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                FillLisDTE(FchDesde, FchHasta, TipoDoc);
            }
        }

        private void FillLisDTE(string FchDesde, string FchHasta, string TipoDoc)
        {
            try
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Llenando grilla. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.DataTable odbtable = null;
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

                odbtable = UIAPIRawForm.DataSources.DataTables.Item("LDTE");
                odbtable.ExecuteQuery(Query);
                Matrix0.Clear();

                #region Matriz
                Matrix0.Columns.Item("RutEmisor").DataBind.Bind("LDTE", "U_RutEmisor");
                Matrix0.Columns.Item("TipoDTE").DataBind.Bind("LDTE", "U_TipoDTE");
                Matrix0.Columns.Item("Folio").DataBind.Bind("LDTE", "U_Folio");
                Matrix0.Columns.Item("Check").DataBind.Bind("LDTE", "Check");
                Matrix0.Columns.Item("RznSoc").DataBind.Bind("LDTE", "U_RznSoc");
                Matrix0.Columns.Item("ExiEmisor").DataBind.Bind("LDTE", "U_ExiEmisor");
                Matrix0.Columns.Item("FchEmis").DataBind.Bind("LDTE", "U_FchEmis");
                Matrix0.Columns.Item("FchVenc").DataBind.Bind("LDTE", "U_FchVenc");
                Matrix0.Columns.Item("FmaPago").DataBind.Bind("LDTE", "U_FmaPago");
                Matrix0.Columns.Item("MntNeto").DataBind.Bind("LDTE", "U_MntNeto");
                Matrix0.Columns.Item("MntExe").DataBind.Bind("LDTE", "U_MntExe");
                Matrix0.Columns.Item("TasaIVA").DataBind.Bind("LDTE", "U_TasaIVA");
                Matrix0.Columns.Item("IVA").DataBind.Bind("LDTE", "U_IVA");
                Matrix0.Columns.Item("MntTotal").DataBind.Bind("LDTE", "U_MntTotal");
                Matrix0.Columns.Item("Glosa").DataBind.Bind("LDTE", "Glosa");
                Matrix0.Columns.Item("FolioRefOC").DataBind.Bind("LDTE", "U_FolioRefOC");
                Matrix0.Columns.Item("FolSAPOC").DataBind.Bind("LDTE", "U_FolioSAPOC");
                Matrix0.Columns.Item("MTotalOC").DataBind.Bind("LDTE", "MTotalOC");
                Matrix0.Columns.Item("FolioRefEM").DataBind.Bind("LDTE", "U_FolioRefEM");
                Matrix0.Columns.Item("FolSAPEM").DataBind.Bind("LDTE", "U_FolioSAPEM");
                Matrix0.Columns.Item("MTotalEM").DataBind.Bind("LDTE", "MTotalEM");
                Matrix0.Columns.Item("FolioRefFA").DataBind.Bind("LDTE", "U_FolioRefFA");
                Matrix0.Columns.Item("FolSAPFA").DataBind.Bind("LDTE", "U_FolioSAPFA");
                Matrix0.Columns.Item("CodRefNC").DataBind.Bind("LDTE", "CodRefNC");
                Matrix0.Columns.Item("Mensaje").DataBind.Bind("LDTE", "Mensaje");
                Matrix0.Columns.Item("DocEntryS").DataBind.Bind("LDTE", "U_DocEntryS");
                Matrix0.Columns.Item("ObjType").DataBind.Bind("LDTE", "U_ObjType");
                Matrix0.Columns.Item("XML").DataBind.Bind("LDTE", "U_XML");
                Matrix0.Columns.Item("PDF64").DataBind.Bind("LDTE", "U_PDF64");
                Matrix0.Columns.Item("Code").DataBind.Bind("LDTE", "Code");
                Matrix0.Columns.Item("CardCode").DataBind.Bind("LDTE", "CardCode");
                #endregion Matriz

                //columnas e items invisibles
                Matrix0.Columns.Item("XML").Visible = false;
                Matrix0.Columns.Item("PDF64").Visible = false;
                Matrix0.Columns.Item("Code").Visible = false;

                Matrix0.LoadFromDataSource();
                FindRef();
                if (Matrix0.RowCount > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documentos encontrados con exito.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else if (Matrix0.RowCount == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("No existen documento con los filtros seleccionados.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("FillLisDTE : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        
        private void FindRef()
        {
            ResultRefSAP resultRefSAP = new ResultRefSAP();
            DTENewSign dTENewSign = new DTENewSign();
            SAPbouiCOM.ProgressBar progressBar;
            progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Buscando referencias por documentos. por favor espere.!", Matrix0.RowCount, true);
            try
            {
                DTECompany dTECompany = dTENewSign.getConfig();
                SAPbouiCOM.DataTable dtable = null;
                dtable = UIAPIRawForm.DataSources.DataTables.Item("LDTE");
                string refOC = null, refEM = null, refFA = null, RutEmisor = null, DocEntryS = null, XML= null;
                double MntTotal = 0;
                string[] folios;
                int indexOfText = 0, index = 0 , integrados =0, noIntegrados = 0;
                String matrixSchemaXML = null, subs = null;
                matrixSchemaXML = (this.Matrix0).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                //SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Buscando referencias por documentos. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                
                progressBar.Text = "Buscando referencias.!";
                for (int i = 1; i <= Matrix0.RowCount; i++)
                {
                    progressBar.Value = i;
                    RutEmisor = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("RutEmisor").Cells.Item(i).Specific).String;
                    XML = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(i).Specific).String;
                    
                    //MntTotal = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntTotal").Cells.Item(i).Specific).String);
                   
                    DocEntryS = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DocEntryS").Cells.Item(i).Specific).String;
                    refOC = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolioRefOC").Cells.Item(i).Specific).String;
                    refEM = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolioRefEM").Cells.Item(i).Specific).String;
                    refFA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolioRefFA").Cells.Item(i).Specific).String;

                    #region Con referencia a OC
                    if (!string.IsNullOrEmpty(refOC) && string.IsNullOrEmpty(DocEntryS))
                    {
                        folios = refOC.Split(',')
                                       .Select(folRef=> "'" + folRef.Trim() +"'")
                                       .ToArray();
                        refOC = (string.Join(",", folios.ToList()));
                        //refOC = "'" + (string.Join("','", folios.ToList())) + "'";
                        resultRefSAP = common.GetRefOC(refOC, RutEmisor, dTECompany.LKPRCOR);
                        if (resultRefSAP.existe)
                        {
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPOC").Cells.Item(i).Specific).String = resultRefSAP.DocNum;
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MTotalOC").Cells.Item(i).Specific).Value = resultRefSAP.DocTotal.ToString();

                            //indexOfText = matrixSchemaXML.IndexOf("<UniqueID>MTotalOC</UniqueID>");
                            //subs = matrixSchemaXML.Substring(0, indexOfText);
                            //index = System.Text.RegularExpressions.Regex.Matches(subs, "<ColumnInfo>").Count - 1;

                            //if (resultRefSAP.DocTotal == MntTotal)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.greenForeColor);
                            //}
                            //else if (resultRefSAP.DocTotal == 0)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.redBackColor);
                            //}
                            //else
                            //{

                            //}
                        }
                    }
                    #endregion  Con referencia a OC

                    #region Con referencia a EM
                    if (!string.IsNullOrEmpty(refEM) && string.IsNullOrEmpty(DocEntryS))
                    {
                        folios = refEM.Split(',')
                                       .Select(folRef => "'" + folRef.Trim() + "'")
                                       .ToArray();
                        refEM = (string.Join(",", folios.ToList()));
                        //refEM = "'" + (string.Join("','", folios.ToList())) + "'";
                        resultRefSAP = common.GetRefEM(refEM, RutEmisor, dTECompany.LKPRCDN);
                        if (resultRefSAP.existe)
                        {
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPEM").Cells.Item(i).Specific).String = resultRefSAP.DocNum;
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MTotalEM").Cells.Item(i).Specific).Value = resultRefSAP.DocTotal.ToString();
                            //Matrix0.SetCellWithoutValidation(i, "FolSAPEM", resultRefSAP.DocNum.ToString());
                            //dtable.SetValue("MTotalEM", i - 1, resultRefSAP.DocTotal.ToString());
                            //dtable.SetValue("FolSAPEM", i - 1, resultRefSAP.DocNum.ToString());
                            
                            //indexOfText = matrixSchemaXML.IndexOf("<UniqueID>MTotalEM</UniqueID>");
                            //subs = matrixSchemaXML.Substring(0, indexOfText);
                            //index = System.Text.RegularExpressions.Regex.Matches(subs, "<ColumnInfo>").Count - 1;

                            //if (resultRefSAP.DocTotal == MntTotal)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.greenForeColor);
                            //}
                            //else if (resultRefSAP.DocTotal == 0)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.redBackColor);
                            //}
                            //else
                            //{

                            //}
                        }
                    }
                    #endregion  Con referencia a EM

                    #region Con referencia a FA
                    if (!string.IsNullOrEmpty(refFA) && string.IsNullOrEmpty(DocEntryS))
                    {
                        folios = refFA.Split(',')
                                      .Select(folRef => "'" + folRef.Trim() + "'")
                                      .ToArray();
                        refFA = (string.Join(",", folios.ToList()));

                        #region CodRef
                        var DTENC = dTENewSign.ObtenerDTE(XML);
                        var RefFA = DTENC.DTE.Referencia.Where(j => j.TpoDocRef == "33" || j.TpoDocRef == "34").ToList();
                        var CodRef = string.Join(", ", RefFA.Select(z => z.CodRef));
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CodRefNC").Cells.Item(i).Specific).String = CodRef;

                        //1: Anula Documento de Referencia
                        //2: Corrige Texto Documento de Referencia
                        //3: Corrige montos
                        #endregion CodRef

                        resultRefSAP = common.GetRefFA(refFA, RutEmisor);
                        if (resultRefSAP.existe)
                        {
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(i).Specific).String = resultRefSAP.DocNum;

                            
                            
                            //((SAPbouiCOM.EditText)Matrix0.Columns.Item("MTotalFA").Cells.Item(i).Specific).Value = resultRefSAP.DocTotal.ToString();
                            //Matrix0.SetCellWithoutValidation(i, "FolSAPEM", resultRefSAP.DocNum.ToString());
                            //dtable.SetValue("MTotalEM", i - 1, resultRefSAP.DocTotal.ToString());
                            //dtable.SetValue("FolSAPEM", i - 1, resultRefSAP.DocNum.ToString());


                            //indexOfText = matrixSchemaXML.IndexOf("<UniqueID>MTotalEM</UniqueID>");
                            //subs = matrixSchemaXML.Substring(0, indexOfText);
                            //index = System.Text.RegularExpressions.Regex.Matches(subs, "<ColumnInfo>").Count - 1;

                            //if (resultRefSAP.DocTotal == MntTotal)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.greenForeColor);
                            //}
                            //else if (resultRefSAP.DocTotal == 0)
                            //{
                            //    Matrix0.CommonSetting.SetCellBackColor(i, index, common.redBackColor);
                            //}
                            //else
                            //{

                            //}
                        }
                    }
                    #endregion  Con referencia a FA

                    #region Documento integrado
                    if (!string.IsNullOrEmpty(DocEntryS)) //Documento integrado
                    {
                        //StatusInt(i);
                        integrados++;
                    }
                    #endregion Documento integrado

                    #region Sin XML
                    if (string.IsNullOrEmpty(XML))
                    {
                        SinXML(i);
                    }
                    #endregion Sin XML

                    #region Totales 
                    noIntegrados = Matrix0.RowCount - integrados;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TotalInt").Specific).String = Convert.ToString(integrados, cultureInfo);
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TotalNInt").Specific).String = Convert.ToString(noIntegrados, cultureInfo);
                    #endregion Totales

                    Matrix0.FlushToDataSource();
                }
                //Matrix0.LoadFromDataSource();
                //Matrix0.AutoResizeColumns();
                //this.UIAPIRawForm.Refresh();
            }
            catch(Exception ex)
            {
                progressBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(progressBar);
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("FindRef : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                progressBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(progressBar);
                progressBar = null;
            }
        }

        private void StatusInt(int Row)
        {
            string DocEntryS = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DocEntryS").Cells.Item(Row).Specific).String;
            
            if (!string.IsNullOrEmpty(DocEntryS))
            {
                //int indexOfText = 0, index = 0;
                //String matrixSchemaXML = null, subs = null;
                //matrixSchemaXML = (this.Matrix0).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);

                //indexOfText = matrixSchemaXML.IndexOf("<UniqueID>Check</UniqueID>");
                //subs = matrixSchemaXML.Substring(0, indexOfText);
                //index = System.Text.RegularExpressions.Regex.Matches(subs, "<ColumnInfo>").Count - 1;

                //Matrix0.CommonSetting.SetCellEditable(Row, index, false);
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Mensaje").Cells.Item(Row).Specific).String = "Integrado";
            }
        }

        private void SinXML(int Row)
        {
            string XML = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(Row).Specific).String;
            if (string.IsNullOrEmpty(XML))
            {
                int indexOfText = 0, index = 0;
                String matrixSchemaXML = null, subs = null;
                matrixSchemaXML = (this.Matrix0).SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                indexOfText = matrixSchemaXML.IndexOf("<UniqueID>Mensaje</UniqueID>");
                subs = matrixSchemaXML.Substring(0, indexOfText);
                index = System.Text.RegularExpressions.Regex.Matches(subs, "<ColumnInfo>").Count - 1;
                Matrix0.CommonSetting.SetCellBackColor(Row, index, common.redBackColor);
            }
        }
        private void FillBP(int Row)
        {
            GC.Collect();
            try
            {
                if(Row>0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("LLenando proveedor. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    bool ExiEmisor = false;

                    ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(Row).Specific).Checked;
                    if (!ExiEmisor)
                    {
                        string decodeStringXml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(Row).Specific).Value;

                        if (!string.IsNullOrEmpty(decodeStringXml))
                        {
                            Common.DTE.ResultDTE resultDTE = new Common.DTE.ResultDTE();
                            DTERECEP.Common.DTENewSign dTENewSign = new Common.DTENewSign();
                            resultDTE = dTENewSign.ObtenerDTE(decodeStringXml);

                            SAPbouiCOM.Framework.Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_BusinessPartner, String.Empty, String.Empty);
                            SAPbouiCOM.Form oFormSN = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

                            string CardCode = null;
                            CardCode = String.Format("P{0}", resultDTE.DTE.Emisor.RUTEmisor.Replace("-", ""));

                            oFormSN.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("7").Specific).Value = resultDTE.DTE.Emisor.RznSoc;
                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("128").Specific).Value = resultDTE.DTE.Emisor.RznSoc;
                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oFormSN.Items.Item("40").Specific;
                            oCombo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oCombo = (SAPbouiCOM.ComboBox)oFormSN.Items.Item("16").Specific;
                            oCombo.Select("101", SAPbouiCOM.BoSearchKey.psk_ByValue);

                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("2014").Specific).Value = resultDTE.DTE.Emisor.GiroEmis;

                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("41").Specific).Value = resultDTE.DTE.Emisor.RUTEmisor;
                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("43").Specific).Value = resultDTE.DTE.Emisor.Telefono;
                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("60").Specific).Value = resultDTE.DTE.Emisor.CorreoEmisor;
                            ((SAPbouiCOM.EditText)oFormSN.Items.Item("5").Specific).Value = CardCode;

                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Proveedor fue llenado con exito.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        }
                        else
                        {
                                 SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("XML no encontrado",1,"Ok");
                        }
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("El proveedor seleccionado ya existe en SAP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }

            }
            catch(Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("FillBP : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Con_OC(int Row)
        {
            try
            {
                if (Row > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Integrando con orden de compra. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    bool ExiEmisor = false;

                    ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(Row).Specific).Checked;
                    if (ExiEmisor)
                    {
                        int iError = 0;
                        string sError = null, DocEntryS = null;
                        string decodeStringXml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(Row).Specific).Value;
                        string RutEmisor = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("RutEmisor").Cells.Item(Row).Specific).String;
                        string TipoDTE = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TipoDTE").Cells.Item(Row).Specific).Value;
                        string Folio = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Folio").Cells.Item(Row).Specific).Value;
                        string FolSAPOC = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPOC").Cells.Item(Row).Specific).Value;
                        Double MntTotal = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntTotal").Cells.Item(Row).Specific).String);
                        string MntNeto = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntNeto").Cells.Item(Row).Specific).Value;
                        string MntExe = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntExe").Cells.Item(Row).Specific).Value;
                        string IVA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("IVA").Cells.Item(Row).Specific).Value;
                        string Code = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Code").Cells.Item(Row).Specific).Value;
                        string CardCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(Row).Specific).Value;
                        string TaxDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchEmis").Cells.Item(Row).Specific).String;
                        string DocDueDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchVenc").Cells.Item(Row).Specific).String;
                        if (decodeStringXml != null)
                        {
                            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
                            string[] folios = FolSAPOC.Split(',');
                            FolSAPOC = "'" + (string.Join("','", folios.ToList())) + "'";

                            docDetSAPs = common.GetDtRefOC(FolSAPOC, RutEmisor);
                            SAPbobsCOM.Documents oDoc = null;

                            switch (TipoDTE)
                            {
                                case "33":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                    break;
                                case "34":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                    break;
                                case "56":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo;
                                    break;
                                case "61":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;
                                    break;
                            }
                            #region header
                            oDoc.CardCode = CardCode;
                            oDoc.DocDate = DateTime.Now;
                            if (!string.IsNullOrEmpty(DocDueDate))
                            {
                                oDoc.DocDueDate = DateTime.ParseExact(DocDueDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            oDoc.TaxDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            oDoc.DocDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            oDoc.FolioPrefixString = TipoDTE;
                            oDoc.FolioNumber = int.Parse(Folio);
                            oDoc.Indicator = TipoDTE;

                            #region tipo factura compra
                            if (ComboBox1.Selected.Value.Equals("Y"))
                            {
                                oDoc.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES;
                            }
                            else
                            {
                                oDoc.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tNO;
                            }
                            #endregion tipo factura compra

                            string DocType = docDetSAPs.Select(x => x.DocType).First().ToString();
                            if (DocType.Equals("I"))
                            {
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                            }
                            else
                            {
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                            }

                            #endregion header

                            #region Lines
                            foreach (DocDetSAP docDetSAP in docDetSAPs)
                            {
                                oDoc.Lines.BaseEntry = docDetSAP.DocEntry;
                                oDoc.Lines.BaseLine = docDetSAP.LineNum;
                                oDoc.Lines.BaseType = int.Parse(docDetSAP.ObjType);
                                oDoc.Lines.Add();

                            }
                            #endregion Lines

                            #region footer 
                            oDoc.DocTotal = MntTotal;
                            #endregion footer
                            iError = oDoc.Add();
                            Common.CRUD.CRUD_ASRDTE cRUD_ASRDTE = new Common.CRUD.CRUD_ASRDTE();
                            if (iError == 0)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documento integrado con exito.! Folio : " +Folio , SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                Conex.oCompany.GetNewObjectCode(out DocEntryS);
                                cRUD_ASRDTE.UpdateASRDTE(Code, DocEntryS, oDoc.DocObjectCode.ToString());                                
                            }
                            else
                            {
                                sError = Conex.oCompany.GetLastErrorDescription();
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error : " +sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        else
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("No se pudo recuperar la información!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("El proveedor del documento seleccionado, no existe en SAP.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }

            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Con_OC : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Con_EM(int Row)
        {
            try
            {
                if (Row > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Integrando con entrada de mercancia. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    bool ExiEmisor = false;
                    ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(Row).Specific).Checked;
                    if (ExiEmisor)
                    {
                        int iError = 0;
                        string sError = null, DocEntryS = null;
                        string decodeStringXml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(Row).Specific).Value;
                        string RutEmisor = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("RutEmisor").Cells.Item(Row).Specific).String;
                        string TipoDTE = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TipoDTE").Cells.Item(Row).Specific).Value;
                        string Folio = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Folio").Cells.Item(Row).Specific).Value;
                        string FolSAPEM = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPEM").Cells.Item(Row).Specific).Value;
                        Double MntTotal =Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntTotal").Cells.Item(Row).Specific).String);
                        string MntNeto = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntNeto").Cells.Item(Row).Specific).Value;
                        string MntExe = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntExe").Cells.Item(Row).Specific).Value;
                        string IVA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("IVA").Cells.Item(Row).Specific).Value;
                        string Code = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Code").Cells.Item(Row).Specific).Value;
                        string CardCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(Row).Specific).Value;
                        string TaxDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchEmis").Cells.Item(Row).Specific).String;
                        string DocDueDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchVenc").Cells.Item(Row).Specific).String;

                        if (decodeStringXml != null)
                        {
                            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
                            string[] folios = FolSAPEM.Split(',');
                            FolSAPEM = "'" + (string.Join("','", folios.ToList())) + "'";

                            docDetSAPs = common.GetDtRefEM(FolSAPEM, RutEmisor);
                            SAPbobsCOM.Documents oDoc = null;

                            switch (TipoDTE)
                            {
                                case "33":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                    break;
                                case "34":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
                                    break;
                                case "56":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                                    oDoc.DocumentSubType = SAPbobsCOM.BoDocumentSubType.bod_PurchaseDebitMemo;
                                    break;                              
                            }
                            #region header
                            oDoc.CardCode = CardCode;
                            //oDoc.DocDate = DateTime.Now;
                            if (!string.IsNullOrEmpty(DocDueDate))
                            {
                                oDoc.DocDueDate = DateTime.ParseExact(DocDueDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            oDoc.TaxDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            oDoc.DocDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            oDoc.FolioPrefixString = TipoDTE;
                            oDoc.FolioNumber = int.Parse(Folio);
                            oDoc.Indicator = TipoDTE;

                            string DocType = docDetSAPs.Select(x => x.DocType).First().ToString();
                            if (DocType.Equals("I"))
                            {
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                            }
                            else
                            {
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                            }

                            #endregion header

                            #region Lines
                            foreach (DocDetSAP docDetSAP in docDetSAPs)
                            {
                                oDoc.Lines.BaseEntry = docDetSAP.DocEntry;
                                oDoc.Lines.BaseLine = docDetSAP.LineNum;
                                oDoc.Lines.BaseType = int.Parse(docDetSAP.ObjType);
                                oDoc.Lines.Add();
                            }
                            #endregion Lines

                            #region footer 
                            oDoc.DocTotal = MntTotal;
                            #endregion footer
                            iError = oDoc.Add();
                            Common.CRUD.CRUD_ASRDTE cRUD_ASRDTE = new Common.CRUD.CRUD_ASRDTE();
                            if (iError == 0)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documento integrado con exito.! Folio : " + Folio, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                Conex.oCompany.GetNewObjectCode(out DocEntryS);
                                cRUD_ASRDTE.UpdateASRDTE(Code, DocEntryS, oDoc.DocObjectCode.ToString());
                            }
                            else
                            {
                                sError = Conex.oCompany.GetLastErrorDescription();
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error : " + sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        else
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("No se pudo recuperar la información.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("El proveedor del documento seleccionado, no existe en SAP.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }

            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Con_EM : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Int_NC(int Row)
        {
            try
            {
                if (Row > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Integrando nota de crédito. por favor espere.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    bool ExiEmisor = false;
                    ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(Row).Specific).Checked;
                    if (ExiEmisor)
                    {
                        int iError = 0;
                        string sError = null, DocEntryS = null;
                        string decodeStringXml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(Row).Specific).Value;
                        string RutEmisor = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("RutEmisor").Cells.Item(Row).Specific).String;
                        string TipoDTE = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TipoDTE").Cells.Item(Row).Specific).Value;
                        string Folio = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Folio").Cells.Item(Row).Specific).Value;
                        string FolSAPFA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(Row).Specific).Value;
                        Double MntTotal = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntTotal").Cells.Item(Row).Specific).String);
                        int CodRefNC = Convert.ToInt32(((SAPbouiCOM.EditText)Matrix0.Columns.Item("CodRefNC").Cells.Item(Row).Specific).String);
                        string MntNeto = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntNeto").Cells.Item(Row).Specific).Value;
                        string MntExe = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("MntExe").Cells.Item(Row).Specific).Value;
                        string IVA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("IVA").Cells.Item(Row).Specific).Value;
                        string Code = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Code").Cells.Item(Row).Specific).Value;
                        string CardCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("CardCode").Cells.Item(Row).Specific).Value;
                        string TaxDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchEmis").Cells.Item(Row).Specific).Value;
                        string DocDueDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchVenc").Cells.Item(Row).Specific).Value;

                        if (decodeStringXml != null)
                        {
                            List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
                            string[] folios = FolSAPFA.Split(',');
                            FolSAPFA = "'" + (string.Join("','", folios.ToList())) + "'";

                            docDetSAPs = common.GetDtRefFA(FolSAPFA, RutEmisor);
                            SAPbobsCOM.Documents oDoc = null;

                            switch (TipoDTE)
                            {
                                case "61":
                                    oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes;                               
                                    break;
                            }
                            #region header
                            oDoc.CardCode = CardCode;
                            oDoc.DocDate = DateTime.Now;
                            if (!string.IsNullOrEmpty(DocDueDate))
                            {
                                oDoc.DocDueDate = DateTime.ParseExact(DocDueDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            }
                            oDoc.TaxDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                            oDoc.FolioPrefixString = TipoDTE;
                            oDoc.FolioNumber = int.Parse(Folio);
                            oDoc.Indicator = TipoDTE;
                            #endregion header

                            #region Lines
                            oDoc.DocumentReferences.ReferencedDocEntry = docDetSAPs[0].DocEntry;
                            oDoc.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_PurchaseInvoice;

                            switch (CodRefNC)
                            {
                                case 1: //1: Anula Documento de Referencia
                                    oDoc.DocumentReferences.Remark = "Anula Documento de Referencia";
                                    foreach (DocDetSAP docDetSAP in docDetSAPs)
                                    {
                                        oDoc.Lines.BaseEntry = docDetSAP.DocEntry;
                                        oDoc.Lines.BaseLine = docDetSAP.LineNum;
                                        oDoc.Lines.BaseType = int.Parse(docDetSAP.ObjType);
                                        oDoc.Lines.Add();
                                    }
                                    break;
                                case 2: //2: Corrige Texto Documento de Referencia
                                    oDoc.DocumentReferences.Remark = "Corrige Texto Documento de Referencia";
                                    break;
                                case 3: //3: Corrige montos
                                    oDoc.DocumentReferences.Remark = "Corrige montos";
                                    break;
                            }
                            
                           
                            #endregion Lines

                            #region footer 
                            oDoc.DocTotal = MntTotal;
                            #endregion footer
                            iError = oDoc.Add();
                            Common.CRUD.CRUD_ASRDTE cRUD_ASRDTE = new Common.CRUD.CRUD_ASRDTE();
                            if (iError == 0)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documento integrado con exito.! Folio : " + Folio, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                Conex.oCompany.GetNewObjectCode(out DocEntryS);
                                cRUD_ASRDTE.UpdateASRDTE(Code, DocEntryS, oDoc.DocObjectCode.ToString());
                            }
                            else
                            {
                                sError = Conex.oCompany.GetLastErrorDescription();
                                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error : " + sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        else
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("No se pudo recuperar la información.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                    else
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("El proveedor del documento seleccionado, no existe en SAP.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    }
                }

            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Con_EM : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void VerPDF(int Row)
        {

            string RUTEmisor = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("RutEmisor").Cells.Item(Row).Specific).Value;
            string TipoDTE = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("TipoDTE").Cells.Item(Row).Specific).Value;
            string Folio = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("Folio").Cells.Item(Row).Specific).Value;
            String Base64 = null;
            string strRutaDestino = System.Environment.CurrentDirectory + @"\" + RUTEmisor + "_" + TipoDTE + "_" + Folio + ".pdf";//  /*System.AppDomain.CurrentDomain.BaseDirectory +*/ @"C:\FEL\TEST\PDF" + RUTEmisor + "_" + TipoDTE + "_" + Folio + ".pdf";
            try
            {
                DTENewSign dTENewSign = new DTENewSign();
                Base64 = dTENewSign.GETPDFDTE(RUTEmisor, TipoDTE, Folio);

                byte[] bytes = Convert.FromBase64String(Base64);
                System.IO.FileStream stream = new System.IO.FileStream(strRutaDestino, System.IO.FileMode.CreateNew);
                System.IO.BinaryWriter writer = new System.IO.BinaryWriter(stream);
                writer.Write(bytes, 0, bytes.Length);
                writer.Close();
            }
            catch //(Exception ex)
            {

            }

            //por tipo de documento igual

            System.Diagnostics.Process.Start(strRutaDestino);// /*System.AppDomain.CurrentDomain.BaseDirectory +*/ @"C:\FEL\TEST\PDF" + rutPdf + "_" + folPdf + ".pdf");

            System.Threading.Thread.Sleep(10000);


            bool result = System.IO.File.Exists(strRutaDestino); ///*System.AppDomain.CurrentDomain.BaseDirectory +*/ @"C:\FEL\TEST\PDF" + rutPdf + "_" + folPdf + ".pdf");
            if (result == true)
            {
                System.IO.File.Delete(strRutaDestino);///*System.AppDomain.CurrentDomain.BaseDirectory +*/ @"C:\FEL\TEST\PDF" + rutPdf + "_" + folPdf + ".pdf");
            }
        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.ButtonCombo ButtonCombo0;

        /// <summary>
        /// Opcion button Process
        /// </summary>
        /// <param name="sboObject"></param>
        /// <param name="pVal"></param>
        private void ButtonCombo0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {


                if (pVal.ActionSuccess && ButtonCombo0.Selected != null)
                {
                    string opcion = ButtonCombo0.Selected.Value;
                    string DocEntryS = null;
                    switch (opcion)
                    {
                        #region Integrar con orden de compra
                        case "Con_OC": //Integrar con orden de compra
                            if (!ComboBox1.Item.Visible)
                            {
                                ComboBox1.Item.Visible = true;
                            }
                            if (ComboBox1.Selected == null && ComboBox1.Item.Visible)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar el tipo de documento destino.!");
                            }
                            else
                            {
                                SAPbouiCOM.CheckBox integrarOC;

                                int cantSelectOC = 0;
                                for (int i = this.Matrix0.RowCount; i > 0; i--)
                                {
                                    integrarOC = (SAPbouiCOM.CheckBox)this.Matrix0.Columns.Item("Check").Cells.Item(i).Specific;
                                    if (integrarOC.Checked)
                                    {
                                        DocEntryS = null;
                                        DocEntryS = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("DocEntryS").Cells.Item(i).Specific).Value;
                                        if (!string.IsNullOrEmpty(DocEntryS))
                                        {
                                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El Documento seleccionado ya existe.!");
                                        }
                                        else
                                        {
                                            cantSelectOC++;
                                            Con_OC(i);
                                        }
                                    }
                                }
                                if (cantSelectOC == 0)
                                {
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento.!");
                                }
                            }
                            break;
                        #endregion Integrar con orden de compra

                        #region Integrar con entrada de mercancia
                        case "Con_EM":
                            ComboBox1.Item.Visible = false;
                            SAPbouiCOM.CheckBox integrarEM;
                            int cantSelectEM = 0;
                            for (int i = this.Matrix0.RowCount; i > 0; i--)
                            {
                                integrarEM = (SAPbouiCOM.CheckBox)this.Matrix0.Columns.Item("Check").Cells.Item(i).Specific;
                                if (integrarEM.Checked)
                                {
                                    DocEntryS = null;
                                    DocEntryS = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("DocEntryS").Cells.Item(i).Specific).Value;
                                    if (!string.IsNullOrEmpty(DocEntryS))
                                    {
                                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El Documento seleccionado ya existe.!");
                                    }
                                    else
                                    {
                                        cantSelectEM++;
                                        Con_EM(i);
                                    }
                                }
                            }
                            if (cantSelectEM == 0)
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento.!");
                            }
                            break;
                        #endregion Integrar con entrada de mercancia

                        #region Integrar nota de crédito
                        case "Int_NC":
                        case "Int_ND":
                            ComboBox1.Item.Visible = false;
                            int RowNC = Matrix0.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (RowNC > 0)
                            {
                                DocEntryS = null;
                                DocEntryS = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("DocEntryS").Cells.Item(RowNC).Specific).Value;
                                if (!string.IsNullOrEmpty(DocEntryS))
                                {
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El Documento seleccionado ya existe.!");
                                }
                                else
                                {
                                    string TipoDTE = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TipoDTE").Cells.Item(RowNC).Specific).Value;
                                    if (TipoDTE.Equals("56") || TipoDTE.Equals("61"))
                                    {
                                        string Xml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(RowNC).Specific).Value;
                                        string Code = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Code").Cells.Item(RowNC).Specific).Value;
                                        string Glosa = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Glosa").Cells.Item(RowNC).Specific).Value;
                                        string TaxDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchEmis").Cells.Item(RowNC).Specific).Value;
                                        string DocDueDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchVenc").Cells.Item(RowNC).Specific).Value;
                                        string FolSAPFA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(RowNC).Specific).Value;
                                        int CodRefNC = Convert.ToInt32(((SAPbouiCOM.EditText)Matrix0.Columns.Item("CodRefNC").Cells.Item(RowNC).Specific).String);

                                        if (string.IsNullOrEmpty(FolSAPFA))
                                        {
                                            int opcionNCND = SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El documento referenciado no existe en SAP desea integrar?", 2, "Si!", "No");
                                            if (opcionNCND == 1)
                                            {
                                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Se integrara el documento sin referencia!");
                                                FDocument fDocument = new FDocument(Xml, Code, Glosa, TaxDate, DocDueDate, FolSAPFA, CodRefNC);
                                            }
                                        }
                                        else
                                        {
                                            FDocument fDocument = new FDocument(Xml, Code, Glosa, TaxDate, DocDueDate, FolSAPFA, CodRefNC);
                                        }
                                    }
                                    else
                                    {
                                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento de tipo " + (opcion == "Int_ND" ? "Nota de débito.!" : "nota de crédito.!"));
                                    }
                                }
                            }
                            else
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento.!");
                            }
                            break;
                            #endregion Integrar nota de crédito                           

                        #region Integrar como servicio
                        case "Como_Ser":
                            ComboBox1.Item.Visible = false;
                            int RowSr = Matrix0.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (RowSr > 0)
                            {
                                DocEntryS = null;
                                DocEntryS = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("DocEntryS").Cells.Item(RowSr).Specific).Value;
                                if (!string.IsNullOrEmpty(DocEntryS))
                                {
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El Documento seleccionado ya existe.!");
                                }
                                else
                                {
                                    string TipoDTE = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TipoDTE").Cells.Item(RowSr).Specific).Value;
                                    if (TipoDTE.Equals("33") || TipoDTE.Equals("34"))
                                    {
                                        bool ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(RowSr).Specific).Checked;
                                        if (ExiEmisor)
                                        {
                                            string Xml = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("XML").Cells.Item(RowSr).Specific).Value;
                                            string Code = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Code").Cells.Item(RowSr).Specific).Value;
                                            string Glosa = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Glosa").Cells.Item(RowSr).Specific).Value;
                                            string TaxDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchEmis").Cells.Item(RowSr).Specific).Value;
                                            string DocDueDate = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FchVenc").Cells.Item(RowSr).Specific).Value;
                                            FDocument fDocument = new FDocument(Xml, Code, Glosa, TaxDate, DocDueDate);
                                        }
                                        else
                                        {
                                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El proveedor debe existir.!");
                                        }
                                    }
                                    else
                                    {
                                        SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento de tipo factura.!");
                                    }
                                }
                            }
                            else
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un documento.!");
                            } 
                            break;
                        #endregion Integrar como servicio

                        #region Llenar BP
                        case "Fill_BP": //Llenar BP
                            ComboBox1.Item.Visible = false;
                            int Row = Matrix0.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (Row > 0)
                            {
                                bool ExiEmisor = ((SAPbouiCOM.CheckBox)Matrix0.Columns.Item("ExiEmisor").Cells.Item(Row).Specific).Checked;
                                if (!ExiEmisor)
                                {
                                    FillBP(Row);
                                }
                                else
                                {
                                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("El proveedor seleccionado ya existe.!");
                                }
                            }
                            else
                            {
                                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Debe seleccionar un proveedor.!");
                            }
                            break;
                        #endregion Llenar BP

                        default:
                            ComboBox1.Item.Visible = false;
                            break;
                    }

                }
                else SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Debe seleccionar una opción ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Process : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private SAPbouiCOM.ComboBox ComboBox1;

        private void ComboBox1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

        }

        private void ButtonCombo0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ActionSuccess)
            {
                string opcion = ButtonCombo0.Selected.Value;
                switch (opcion)
                {
                    #region Integrar con orden de compra
                    case "Con_OC": //Integrar con orden de compra
                        if (!ComboBox1.Item.Visible)
                        {
                            ComboBox1.Item.Visible = true;
                        }
                        break;
                        #endregion Integrar con orden de compra

                    case "Fill_BP": 
                    case "Con_EM":
                    case "Como_Ser":
                    case "Int_NC":
                    case "Int_ND":
                        ComboBox1.Item.Visible = false;
                        break;

                    default:
                        ComboBox1.Item.Visible = false;
                        break;
                }
            }
        }

        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.ChooseFromList oCFL = null; SAPbouiCOM.Conditions oCons = null;
            try
            {
                switch (pVal.ColUID)
                {
                    #region Condition purchase order
                    case "FolSAPOC":
                        string CardCode = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("CardCode").Cells.Item(pVal.Row).Specific).Value;
                        if (!string.IsNullOrEmpty(CardCode))
                        {
                            oCFL = this.UIAPIRawForm.ChooseFromLists.Item("CFPOR");
                            oCons = new SAPbouiCOM.Conditions();
                            oCFL.SetConditions(oCons);
                            SAPbouiCOM.Condition oCon2 = oCons.Add();

                            oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon2.CondVal = "O";
                            oCon2.Alias = "DocStatus";
                            oCFL.SetConditions(oCons);

                            oCon2.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                            oCon2 = oCons.Add();
                            oCon2.Alias = "LicTradNum";
                            oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon2.CondVal = CardCode;

                            oCFL.SetConditions(oCons);
                        }
                        else
                        {
                            SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Socio de negocio no existe SAP.!");
                        }
                        break;
                        #endregion Condition purchase order
                }

            }
            catch //(Exception ex)
            {

            }
        }

        private void SelectChooise(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            /*
            SAPbouiCOM.ISBOChooseFromListEventArg oCFLEvento = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.DataTable oDataTable = null;
            
            try
            {
                oCFLEvento = ((SAPbouiCOM.ISBOChooseFromListEventArg)pVal);
                oCFL = this.UIAPIRawForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                oDataTable = oCFLEvento.SelectedObjects;
                string val = null, val2 = null;
                List<string> valores = new List<string>();

                if (oDataTable != null)
                {
                    try
                    {
                        #region Valores dependiendo del objeto
                        switch (oCFL.ObjectType)
                        {
                            case "22": //OC
                                for (int i = 0; i < oDataTable.Rows.Count; i++)
                                {
                                    val = System.Convert.ToString(oDataTable.GetValue("DocNum", i));
                                    val2 = System.Convert.ToString(oDataTable.GetValue(1, i));
                                    valores.Add(val);
                                }
                                break;
                            case "20": //EM
                                for (int i = 0; i < oDataTable.Rows.Count; i++)
                                {
                                    val = System.Convert.ToString(oDataTable.GetValue("DocNum", i));
                                    val2 = System.Convert.ToString(oDataTable.GetValue(1, i));
                                    valores.Add(val);
                                }
                                break;
                        }
                        #endregion Valores dependiendo del objeto
                    }
                    catch
                    { }

                    double totalOc = 0;
                    double totalEm = 0;
                    int totalDo = 0;
                    string folioocinit = "0";
                    switch (pVal.ColUID)
                    {
                        case "docBaseOC":

                            ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_folOC").Cells.Item(pVal.Row).Specific).Value = string.Join(",", valores);

                            #region Set Values
                            List<string> folios = new List<string>();
                            foreach (string folio in valores)
                            {
                                folios.Add("'" + folio + "'");
                            }
                            string Folio = string.Join(",", folios);
                            if (!string.IsNullOrEmpty(Folio))
                            {
                                string Rut = ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_Rut").Cells.Item(pVal.Row).Specific).Value;
                                double totalDoc = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_MonTot").Cells.Item(pVal.Row).Specific).Value, invC);

                                FuncionesComunes.Refern reff = new FuncionesComunes.Refern();
                                reff = FuncionesComunes.SearchReferxDoc("801", Folio.Replace(";", ","), Rut.Trim(), LKOC, LKEM, "");
                                if (reff.blExiste)
                                {
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiOC", pVal.Row - 1, "Y");
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("folOC", pVal.Row - 1, reff.referencia.Replace(",", ","));
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalOC", pVal.Row - 1, Convert.ToString(reff.monto));
                                    Convert.ToDouble(reff.monto, invC);
                                    totalOc = Convert.ToDouble(reff.monto, invC);
                                    Matrix2.LoadFromDataSource();
                                }
                                else
                                {
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiOC", pVal.Row - 1, "N");
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalOC", pVal.Row - 1, 0);
                                }
                                sDifer = Convert.ToString(totalOc - totalDoc);
                                ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_Difere").Cells.Item(pVal.Row).Specific).Value = sDifer;

                                Matrix2.LoadFromDataSource();
                            }
                            else
                            {
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("folOC", pVal.Row - 1, "");
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiOC", pVal.Row - 1, "N");
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalOC", pVal.Row - 1, 0);

                                Matrix2.LoadFromDataSource();
                            }
                            #endregion Set Values 
                            break;

                        case "docBaseEM":

                            ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_folEM").Cells.Item(pVal.Row).Specific).Value = string.Join(",", valores);

                            #region Set Values 
                            List<string> FolioEM = new List<string>();
                            //SAPbouiCOM.Matrix oGrilla = oForm.Items.Item("oMtx").Specific;
                            foreach (string folio in valores)
                            {
                                FolioEM.Add("'" + folio + "'");
                            }
                            string Folioem = string.Join(",", FolioEM);
                            if (!string.IsNullOrEmpty(Folioem))
                            {
                                string Rut = ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_Rut").Cells.Item(pVal.Row).Specific).Value;
                                double totalDoc = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_MonTot").Cells.Item(pVal.Row).Specific).Value, invC);

                                FuncionesComunes.Refern reff = new FuncionesComunes.Refern();
                                reff = FuncionesComunes.SearchReferxDoc("52", Folioem.Replace(";", ","), Rut.Trim(), LKOC, LKEM, "");
                                if (reff.blExiste)
                                {
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiEM", pVal.Row - 1, "Y");
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("folEM", pVal.Row - 1, reff.referencia.Replace(",", ","));
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalEM", pVal.Row - 1, Convert.ToString(reff.monto));
                                    Convert.ToDouble(reff.monto, invC);
                                    totalEm = Convert.ToDouble(reff.monto, invC);
                                    Matrix2.LoadFromDataSource();
                                }
                                else
                                {
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiEM", pVal.Row - 1, "N");
                                    UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalEM", pVal.Row - 1, 0);
                                }
                                sDifer = Convert.ToString(totalEm - totalDoc);
                                ((SAPbouiCOM.EditText)Matrix2.Columns.Item("Col_DifEM").Cells.Item(pVal.Row).Specific).Value = sDifer;

                                Matrix2.LoadFromDataSource();
                            }
                            else
                            {
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("folEM", pVal.Row - 1, "");
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("exiEM", pVal.Row - 1, "N");
                                UIAPIRawForm.DataSources.DataTables.Item("DTDocDTE").SetValue("totalEM", pVal.Row - 1, 0);

                                Matrix2.LoadFromDataSource();
                            }
                            #endregion Set Values 
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error SelectChooise :" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            */
        }

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;            
        }

        private void Matrix0_LinkPressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            switch (pVal.ColUID)
            {
                case "DocEntryS":
                    string TipoDTE = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("TipoDTE").Cells.Item(pVal.Row).Specific).Value;
                    string DocEntryS = ((SAPbouiCOM.EditText)this.Matrix0.Columns.Item("DocEntryS").Cells.Item(pVal.Row).Specific).Value;
                    switch (TipoDTE)
                    {
                        case "33":
                        case "34":
                        case "56":
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoice, null, DocEntryS);
                            break;
                        case "61":
                            Application.SBO_Application.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_PurchaseInvoiceCreditMemo, null, DocEntryS);
                            break;
                    }
                    break;
                case "PDF":
                    VerPDF(pVal.Row);
                    break;
            }
        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ItemChanged)
                {
                    switch (pVal.ColUID)
                    {
                        case "FolioRefFA":
                            string refFA = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolioRefFA").Cells.Item(pVal.Row).Specific).String;
                            string DocEntryS = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DocEntryS").Cells.Item(pVal.Row).Specific).String;
                            string RutEmisor = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("RutEmisor").Cells.Item(pVal.Row).Specific).String;
                            #region Con referencia a FA
                            string[] folios;
                            if (!string.IsNullOrEmpty(refFA) && string.IsNullOrEmpty(DocEntryS))
                            {
                                folios = refFA.Split(',')
                                    .Select(folRef => "'" + folRef.Trim() + "'")
                                    .ToArray();
                                refFA = (string.Join(",", folios.ToList()));

                                #region CodRef
                                //var DTENC = dTENewSign.ObtenerDTE(XML);
                                //var RefFA = DTENC.DTE.Referencia.Where(j => j.TpoDocRef == "33" || j.TpoDocRef == "34").ToList();
                                //var CodRef = string.Join(", ", RefFA.Select(z => z.CodRef));
                                //((SAPbouiCOM.EditText)Matrix0.Columns.Item("CodRefNC").Cells.Item(i).Specific).String = CodRef;

                                //1: Anula Documento de Referencia
                                //2: Corrige Texto Documento de Referencia
                                //3: Corrige montos
                                #endregion CodRef
                                ResultRefSAP resultRefSAP = new ResultRefSAP();
                                resultRefSAP = common.GetRefFA(refFA, RutEmisor);
                                if (resultRefSAP.existe)
                                {
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(pVal.Row).Specific).String = resultRefSAP.DocNum;
                                }
                                else
                                {
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(pVal.Row).Specific).String = "";
                                }
                                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("FolSAPFA").Cells.Item(pVal.Row).Specific).Active = true;
                            }
                            #endregion  Con referencia a FA
                            break;
                    }
                }
            }
            catch //(Exception ex)
            {

            }
        }
    }
}
