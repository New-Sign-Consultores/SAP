using DTERECEP.Common;
using DTERECEP.Common.DTE;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DTERECEP.Forms
{
    [FormAttribute("DTERECEP.Forms.FDocument", "Forms/FDocument.b1f")]
    class FDocument : UserFormBase
    {
        CultureInfo cultureInfo = new CultureInfo("es-CL");
        Common.Common common = new Common.Common();
        public FDocument()
        {
        }

        public FDocument( string XML, string Code,string Glosa,string TaxDate, string DocDueDate) 
        {
            try
            {
                Common.DTENewSign dTENewSign = new Common.DTENewSign();
                Common.Common common = new Common.Common();
                ResultDTE resultDTE = dTENewSign.ObtenerDTE(XML);

                if (resultDTE.DTE != null)
                {
                    BPSAP bPSAP = common.GetBPSAP(resultDTE.DTE.Emisor.RUTEmisor);

                    #region head
                    
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Code").Specific).String = Code;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific).String = bPSAP.CardCode;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific).String = bPSAP.CardName;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioPref").Specific).String = resultDTE.DTE.IdDoc.TipoDTE.ToString();
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioNum").Specific).String = resultDTE.DTE.IdDoc.Folio.ToString();
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TaxDate").Specific).String = DateTime.ParseExact(TaxDate,"yyyyMMdd",System.Globalization.CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                    if (!string.IsNullOrEmpty(DocDueDate))
                    {
                        ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDueDate").Specific).String = DateTime.ParseExact(DocDueDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                    }
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDate").Specific).String = DateTime.Now.ToString("yyyyMMdd");
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Glosa").Specific).String = Glosa;

                    ((SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("stDocB").Specific).Item.Visible = false;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BaseEntry").Specific).Item.Visible = false;
                    ((SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("CodRef").Specific).Item.Visible = false;

                    ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("DocType").Specific)).Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    this.Matrix1.Item.Visible = false;
                    #endregion head

                    #region Detalle
                    int i = 1;
                    foreach (var detalle in resultDTE.DTE.Detalle)
                    {
                        Matrix0.AddRow();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dscript").Cells.Item(i).Specific).String = detalle.NmbItem;
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("UnitPrice").Cells.Item(i).Specific).String = Convert.ToString(detalle.MontoItem);
                        if(resultDTE.DTE.Totales.TasaIVA > 0)
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA";
                        else
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA_EXE";

                       
                        i++;
                    }
                    #endregion Detalle

                    #region Footer
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MntNeto").Specific).String = Convert.ToString(resultDTE.DTE.Totales.MntNeto);
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("VatSum").Specific).String = Convert.ToString(resultDTE.DTE.Totales.IVA);
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocTotal").Specific).String = Convert.ToString(resultDTE.DTE.Totales.MntTotal);
                    #endregion Footer

                    Matrix0.FlushToDataSource();
                }
                this.Matrix0.AutoResizeColumns();
                this.Matrix1.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("FDocument: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public FDocument(string XML, string Code, string Glosa, string TaxDate, string DocDueDate,string FolSAPFA,int CodRef)
        {
            try
            {
                Common.DTENewSign dTENewSign = new Common.DTENewSign();
                Common.Common common = new Common.Common();
                ResultDTE resultDTE = dTENewSign.ObtenerDTE(XML);

                if (resultDTE.DTE != null)
                {
                    BPSAP bPSAP = common.GetBPSAP(resultDTE.DTE.Emisor.RUTEmisor);

                    #region head

                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Code").Specific).String = Code;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific).String = bPSAP.CardCode;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific).String = bPSAP.CardName;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioPref").Specific).String = resultDTE.DTE.IdDoc.TipoDTE.ToString();
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioNum").Specific).String = resultDTE.DTE.IdDoc.Folio.ToString();
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TaxDate").Specific).String = DateTime.ParseExact(TaxDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                    if (!string.IsNullOrEmpty(DocDueDate))
                    {
                        ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDueDate").Specific).String = DateTime.ParseExact(DocDueDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString("yyyyMMdd");
                    }
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDate").Specific).String = DateTime.Now.ToString("yyyyMMdd");
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Glosa").Specific).String = Glosa;

                    #region Documento base
                    List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
                    string[] folios = FolSAPFA.Split(',');
                    FolSAPFA = "'" + (string.Join("','", folios.ToList())) + "'";

                    docDetSAPs = common.GetDtRefFA(FolSAPFA, resultDTE.DTE.Emisor.RUTEmisor);
                    if(docDetSAPs.Count > 0)
                    {
                        ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BaseEntry").Specific).String = Convert.ToString(docDetSAPs[0].DocEntry);
                    }
                    ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("CodRef").Specific)).Select(CodRef.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                    ((SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("stDocB").Specific).Item.Visible = true;
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BaseEntry").Specific).Item.Visible = true;
                    ((SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("CodRef").Specific).Item.Visible = true;
                    #endregion Documento base

                    #endregion head

                    #region Detalle
                    //Clase de documento
                    ((SAPbouiCOM.StaticText)this.UIAPIRawForm.Items.Item("lbTypeDoc").Specific).Item.Visible = true;
                    ((SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("DocType").Specific).Item.Visible = true;
                    ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("DocType").Specific)).Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    int i = 1;
                    foreach (var detalle in resultDTE.DTE.Detalle)
                    {
                        #region Service
                        Matrix0.AddRow();
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dscript").Cells.Item(i).Specific).String = detalle.NmbItem;
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("UnitPrice").Cells.Item(i).Specific).String = Convert.ToString(detalle.MontoItem);
                        if (resultDTE.DTE.Totales.TasaIVA > 0)
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA";
                        else
                            ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA_EXE";
                        #endregion Service

                        #region Item
                        this.Matrix1.AddRow();
                        ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("Dscript").Cells.Item(i).Specific).String = detalle.NmbItem;
                        ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("Quantity").Cells.Item(i).Specific).String = Convert.ToString(detalle.QtyItem);
                        ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("UnitPrice").Cells.Item(i).Specific).String = Convert.ToString(detalle.MontoItem);
                        if (resultDTE.DTE.Totales.TasaIVA > 0)
                            ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA";
                        else
                            ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("TaxCode").Cells.Item(i).Specific).String = "IVA_EXE";
                        #endregion Item
                        i++;
                    }
                    #endregion Detalle

                    #region Footer
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("MntNeto").Specific).String = Convert.ToString(resultDTE.DTE.Totales.MntNeto);
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("VatSum").Specific).String = Convert.ToString(resultDTE.DTE.Totales.IVA);
                    ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocTotal").Specific).String = Convert.ToString(resultDTE.DTE.Totales.MntTotal);

                    switch (CodRef)
                    {
                        case 1: //1: Anula Documento de Referencia
                            ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Comments").Specific).String = "Anula Documento de Referencia";                           
                            break;
                        case 2: //2: Corrige Texto Documento de Referencia
                            ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Comments").Specific).String = "Corrige Texto Documento de Referencia";
                            break;
                        case 3: //3: Corrige montos
                            ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Comments").Specific).String = "Corrige montos";
                            break;
                    }
                    #endregion Footer

                    Matrix0.FlushToDataSource();
                }
                this.Matrix0.AutoResizeColumns();
                this.Matrix1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("FDocument: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void AddConditionCFL(string CFL, string sCondVal, SAPbouiCOM.BoConditionOperation ConditionOperation, string Alias)
        {
            SAPbouiCOM.ChooseFromList oCFL = null; SAPbouiCOM.Conditions oCons = null;
            try
            {

                oCFL = this.UIAPIRawForm.ChooseFromLists.Item(CFL);
                oCons = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oCons);
                SAPbouiCOM.Condition oCon = oCons.Add();

                oCon.Operation = ConditionOperation; //BoConditionOperation.co_EQUAL
                oCon.CondVal = sCondVal;  // Y
                oCon.Alias = Alias;  //Postable
                oCFL.SetConditions(oCons);
            }
            catch (Exception)
            {
                throw new System.NotImplementedException();
            }
        }
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lbProv").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lbCrName").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lbRefAc").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("CardCode").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("CardName").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("NumAtCard").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_7").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("DocDate").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("TaxDate").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("FolioNum").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("FolioPref").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_17").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("Comments").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("MntNeto").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_22").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("VatSum").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("DocTotal").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("DocNum").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Gion").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtServ").Specific));
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("FIntDTE").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("DocType").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("lbTypeDoc").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("Code").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("stGlosa").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("Glosa").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("DocDueDate").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("stDocB").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("CodRef").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("BaseEntry").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LkBase").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("mtItem").Specific));
            this.Matrix1.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix1_ChooseFromListAfter);
            // this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkSn").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {
            AddConditionCFL("cfACT", "Y", SAPbouiCOM.BoConditionOperation.co_EQUAL, "Postable");
            AddConditionCFL("cfOCR1", "1", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR2", "2", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR3", "3", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR4", "4", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR5", "5", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");

            AddConditionCFL("cfOCR1S", "1", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR2S", "2", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR3S", "3", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR4S", "4", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");
            AddConditionCFL("cfOCR5S", "5", SAPbouiCOM.BoConditionOperation.co_EQUAL, "DimCode");

            this.StaticText3.Item.Visible = false;
            this.ComboBox0.Item.Visible = false;
            this.EditText13.Item.Visible = false;


            this.StaticText14.Item.Visible = false;
            this.ComboBox1.Item.Visible = false;
            this.EditText12.Item.Visible = false;


            this.Matrix0.Columns.Item("VatSum").Visible = false;
            this.Matrix0.Columns.Item("LineTotal").Visible = false;

            
            this.ComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("DocType").Specific)).ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;



            //this.StaticText6.Item.Visible = false;
            //this.EditText4.Item.Visible = false;
        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.LinkedButton LinkedButton0;

        /// <summary>
        /// Funcion agrega el documento
        /// </summary>
        /// <param name="sboObject"></param>
        /// <param name="pVal"></param>
        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string FolioPref = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioPref").Specific).String;
            switch (FolioPref)
            {
                case "33":
                case "34":
                    IntegarSinRef(sboObject, pVal);
                    break;
                case "61":
                case "56":
                    IntegarConRef(sboObject, pVal);
                    break;

            }
        }

        private void IntegarSinRef(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int iError = 0;
            string DocEntryS = null, sError = null;
            try
            {
                string CardCode = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific).String;
                string Code = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Code").Specific).String;
                string CardName = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific).String;
                string FolioPref = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioPref").Specific).String;
                int FolioNum = int.Parse(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioNum").Specific).String);
                string TaxDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TaxDate").Specific).String;
                string DocDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDate").Specific).String;
                string DocDueDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDueDate").Specific).String;
                string NumAtCard = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("NumAtCard").Specific).String;
                string Glosa = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Glosa").Specific).String;
                SAPbobsCOM.Documents oDoc = null;
                switch (FolioPref)
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
                oDoc.DocDate = DateTime.ParseExact(DocDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                if (!string.IsNullOrEmpty(DocDueDate))
                {
                    oDoc.DocDueDate = DateTime.ParseExact(DocDueDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                }
                oDoc.TaxDate = DateTime.ParseExact(TaxDate, "dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                oDoc.FolioPrefixString = FolioPref;
                oDoc.FolioNumber = FolioNum;
                oDoc.Indicator = FolioPref;
                if (!string.IsNullOrEmpty(Glosa))
                {
                    oDoc.JournalMemo = Glosa;
                }
                if (!string.IsNullOrEmpty(NumAtCard)) oDoc.NumAtCard = NumAtCard;
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                #endregion header

                #region Lines
                string Dscript = null, TaxCode = null, AcctCode = null, ActId = null, OcrCode = null, OcrCode2 = null, OcrCode3 = null, OcrCode4 = null, OcrCode5 = null, Project = null;
                double UnitPrice, /*LineTotal,*/ VatSum, DocTotal;
                for (int i = 1; i <= this.Matrix0.RowCount; i++)
                {
                    Dscript = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dscript").Cells.Item(i).Specific).String;
                    UnitPrice = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("UnitPrice").Cells.Item(i).Specific).String, cultureInfo);
                    //LineTotal = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("LineTotal").Cells.Item(i).Specific).String, cultureInfo);
                    TaxCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String;
                    AcctCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("AcctCode").Cells.Item(i).Specific).String;
                    ActId = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("ActId").Cells.Item(i).Specific).String;
                    
                    OcrCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode").Cells.Item(i).Specific).String;
                    OcrCode2 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode2").Cells.Item(i).Specific).String;
                    OcrCode3 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode3").Cells.Item(i).Specific).String;
                    OcrCode4 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode4").Cells.Item(i).Specific).String;
                    OcrCode5 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode5").Cells.Item(i).Specific).String;
                    Project = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Project").Cells.Item(i).Specific).String;

                    oDoc.Lines.ItemDescription = Dscript;
                    oDoc.Lines.UnitPrice = UnitPrice;
                    oDoc.Lines.TaxCode = TaxCode;
                    oDoc.Lines.AccountCode = ActId;

                    if (!string.IsNullOrEmpty(OcrCode)) oDoc.Lines.CostingCode = OcrCode;
                    if (!string.IsNullOrEmpty(OcrCode2)) oDoc.Lines.CostingCode2 = OcrCode2;
                    if (!string.IsNullOrEmpty(OcrCode3)) oDoc.Lines.CostingCode3 = OcrCode3;
                    if (!string.IsNullOrEmpty(OcrCode4)) oDoc.Lines.CostingCode4 = OcrCode4;
                    if (!string.IsNullOrEmpty(OcrCode5)) oDoc.Lines.CostingCode5 = OcrCode5;
                    if (!string.IsNullOrEmpty(Project)) oDoc.Lines.ProjectCode = Project;

                    oDoc.Lines.Add();
                }
                #endregion Lines

                #region Footer                
                string Comments = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Comments").Specific).String;
                VatSum = Convert.ToDouble(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("VatSum").Specific).String, cultureInfo);
                DocTotal = Convert.ToDouble(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocTotal").Specific).String, cultureInfo);
                #endregion Footer

                #region footer 
                if (!string.IsNullOrEmpty(Comments)) oDoc.Comments = Comments;
                oDoc.DocTotal = DocTotal;
                #endregion footer

                iError = oDoc.Add();
                Common.CRUD.CRUD_ASRDTE cRUD_ASRDTE = new Common.CRUD.CRUD_ASRDTE();
                if (iError == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documento integrado con exito.! Folio : " + FolioNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    Conex.oCompany.GetNewObjectCode(out DocEntryS);
                    cRUD_ASRDTE.UpdateASRDTE(Code, DocEntryS, oDoc.DocObjectCode.ToString());
                }
                else
                {
                    sError = Conex.oCompany.GetLastErrorDescription();
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error : " + sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Agregar : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            this.UIAPIRawForm.Close();
        }

        private void IntegarConRef(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int iError = 0;
            string DocEntryS = null, sError = null;
            try
            {
                string CardCode = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardCode").Specific).String;
                string Code = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Code").Specific).String;
                string CardName = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("CardName").Specific).String;
                string FolioPref = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioPref").Specific).String;
                int FolioNum = int.Parse(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("FolioNum").Specific).String);
                string TaxDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("TaxDate").Specific).String;
                string DocDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDate").Specific).String;
                string DocDueDate = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocDueDate").Specific).String;
                string NumAtCard = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("NumAtCard").Specific).String;
                string Glosa = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Glosa").Specific).String;

                string FolSAPFA = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("BaseEntry").Specific).String;
                int CodRef = Convert.ToInt32(((SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("CodRef").Specific).Value);

                SAPbobsCOM.Documents oDoc = null;
                switch (FolioPref)
                {
                    case "56":
                        oDoc = (SAPbobsCOM.Documents)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices;
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

                oDoc.FolioPrefixString = FolioPref;
                oDoc.FolioNumber = FolioNum;
                oDoc.Indicator = FolioPref;
                if (!string.IsNullOrEmpty(Glosa))
                {
                    oDoc.JournalMemo = Glosa;
                }
                if (!string.IsNullOrEmpty(NumAtCard)) oDoc.NumAtCard = NumAtCard;

                #endregion header

                #region DocumentReferences
                //List<DocDetSAP> docDetSAPs = new List<DocDetSAP>();
                //string[] folios = FolSAPFA.Split(',');
                //FolSAPFA = "'" + (string.Join("','", folios.ToList())) + "'";
                //docDetSAPs = common.GetDtRefFA(FolSAPFA, "RutEmisor");
                if (!string.IsNullOrEmpty(FolSAPFA))
                {
                    oDoc.DocumentReferences.ReferencedDocEntry = Convert.ToInt32(FolSAPFA);
                    oDoc.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_PurchaseInvoice;
                    switch (CodRef)
                    {
                        case 1: //1: Anula Documento de Referencia
                            oDoc.DocumentReferences.Remark = "Anula Documento de Referencia";
                            break;
                        case 2: //2: Corrige Texto Documento de Referencia
                            oDoc.DocumentReferences.Remark = "Corrige Texto Documento de Referencia";
                            break;
                        case 3: //3: Corrige montos
                            oDoc.DocumentReferences.Remark = "Corrige montos";
                            break;
                    }
                }
                #endregion DocumentReferences

                #region Lines
                string Dscript = null, TaxCode = null, AcctCode = null, ActId = null, OcrCode = null, OcrCode2 = null, OcrCode3 = null, OcrCode4 = null, OcrCode5 = null, Project = null;
                double UnitPrice, Quantity, VatSum, DocTotal;
                string DocType = ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("DocType").Specific)).Value;
                switch (DocType)
                {
                    #region Service
                    case "S":
                        oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;
                        for (int i = 1; i <= this.Matrix0.RowCount; i++)
                        {
                            Dscript = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dscript").Cells.Item(i).Specific).String;
                            UnitPrice = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix0.Columns.Item("UnitPrice").Cells.Item(i).Specific).String, cultureInfo);
                            TaxCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(i).Specific).String;
                            AcctCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("AcctCode").Cells.Item(i).Specific).String;
                            ActId = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("ActId").Cells.Item(i).Specific).String;
                            OcrCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode").Cells.Item(i).Specific).String;
                            OcrCode2 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode2").Cells.Item(i).Specific).String;
                            OcrCode3 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode3").Cells.Item(i).Specific).String;
                            OcrCode4 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode4").Cells.Item(i).Specific).String;
                            OcrCode5 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode5").Cells.Item(i).Specific).String;
                            Project = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Project").Cells.Item(i).Specific).String;

                            oDoc.Lines.ItemDescription = Dscript;
                            oDoc.Lines.UnitPrice = UnitPrice;
                            oDoc.Lines.TaxCode = TaxCode;
                            oDoc.Lines.AccountCode = ActId;

                            if (!string.IsNullOrEmpty(OcrCode)) oDoc.Lines.CostingCode = OcrCode;
                            if (!string.IsNullOrEmpty(OcrCode2)) oDoc.Lines.CostingCode2 = OcrCode2;
                            if (!string.IsNullOrEmpty(OcrCode3)) oDoc.Lines.CostingCode3 = OcrCode3;
                            if (!string.IsNullOrEmpty(OcrCode4)) oDoc.Lines.CostingCode4 = OcrCode4;
                            if (!string.IsNullOrEmpty(OcrCode5)) oDoc.Lines.CostingCode5 = OcrCode5;
                            if (!string.IsNullOrEmpty(Project)) oDoc.Lines.ProjectCode = Project;

                            oDoc.Lines.Add();
                        }
                        break;
                    #endregion Service

                    #region Item
                    case "I":
                        oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                        for (int i = 1; i <= this.Matrix1.RowCount; i++)
                        {
                            Dscript = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("ItemCode").Cells.Item(i).Specific).String;
                            Quantity = Convert.ToDouble(((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("Quantity").Cells.Item(i).Specific).String, cultureInfo);
                            UnitPrice = Convert.ToDouble(((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("UnitPrice").Cells.Item(i).Specific).String, cultureInfo);
                            TaxCode = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("TaxCode").Cells.Item(i).Specific).String;
                            OcrCode = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode").Cells.Item(i).Specific).String;
                            OcrCode2 = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode2").Cells.Item(i).Specific).String;
                            OcrCode3 = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode3").Cells.Item(i).Specific).String;
                            OcrCode4 = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode4").Cells.Item(i).Specific).String;
                            OcrCode5 = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode5").Cells.Item(i).Specific).String;
                            Project = ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("Project").Cells.Item(i).Specific).String;

                            oDoc.Lines.ItemCode = Dscript;
                            oDoc.Lines.Quantity = Quantity;
                            oDoc.Lines.UnitPrice = UnitPrice;
                            oDoc.Lines.TaxCode = TaxCode;

                            if (!string.IsNullOrEmpty(OcrCode)) oDoc.Lines.CostingCode = OcrCode;
                            if (!string.IsNullOrEmpty(OcrCode2)) oDoc.Lines.CostingCode2 = OcrCode2;
                            if (!string.IsNullOrEmpty(OcrCode3)) oDoc.Lines.CostingCode3 = OcrCode3;
                            if (!string.IsNullOrEmpty(OcrCode4)) oDoc.Lines.CostingCode4 = OcrCode4;
                            if (!string.IsNullOrEmpty(OcrCode5)) oDoc.Lines.CostingCode5 = OcrCode5;
                            if (!string.IsNullOrEmpty(Project)) oDoc.Lines.ProjectCode = Project;

                            oDoc.Lines.Add();
                        }
                        break;
                        #endregion Item
                }
                #endregion Lines

                #region Footer                
                string Comments = ((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("Comments").Specific).String;
                VatSum = Convert.ToDouble(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("VatSum").Specific).String, cultureInfo);
                DocTotal = Convert.ToDouble(((SAPbouiCOM.EditText)this.UIAPIRawForm.Items.Item("DocTotal").Specific).String, cultureInfo);
                #endregion Footer

                #region footer 
                if (!string.IsNullOrEmpty(Comments)) oDoc.Comments = Comments;
                oDoc.DocTotal = DocTotal;
                #endregion footer

                iError = oDoc.Add();
                Common.CRUD.CRUD_ASRDTE cRUD_ASRDTE = new Common.CRUD.CRUD_ASRDTE();
                if (iError == 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Documento integrado con exito.! Folio : " + FolioNum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    Conex.oCompany.GetNewObjectCode(out DocEntryS);
                    cRUD_ASRDTE.UpdateASRDTE(Code, DocEntryS, oDoc.DocObjectCode.ToString());
                }
                else
                {
                    sError = Conex.oCompany.GetLastErrorDescription();
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error : " + sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Agregar : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            this.UIAPIRawForm.Close();
        }
        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SelectChooise(pVal, sboObject);

        }
        private void SelectChooise(SAPbouiCOM.SBOItemEventArg pVal, object sboObject)
        {
            SAPbouiCOM.ISBOChooseFromListEventArg oCFLEvento = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.DataTable oDataTable = null;
            try
            {
                oCFLEvento = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                oCFL = this.UIAPIRawForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                oDataTable = oCFLEvento.SelectedObjects;
                string val = null, val2 = null, AcctCode = null;
                if (oDataTable != null)
                {
                    try
                    {
                        #region Valores dependiendo del objeto
                        switch (oCFL.ObjectType)
                        {
                            default:
                                val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                val2 = System.Convert.ToString(oDataTable.GetValue(1, 0));
                                break;                            
                            case "1":
                                val = System.Convert.ToString(oDataTable.GetValue("ActId", 0));
                                AcctCode = System.Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                                val2 = System.Convert.ToString(oDataTable.GetValue("AcctName", 0));
                                break;
                            case "4":
                                val = System.Convert.ToString(oDataTable.GetValue("ItemCode", 0));
                                val2 = System.Convert.ToString(oDataTable.GetValue("ItemName", 0));
                                break;
                            case "128":
                                val = System.Convert.ToString(oDataTable.GetValue("Code", 0));
                                val2 = System.Convert.ToString(oDataTable.GetValue("Rate", 0));
                                break;
                        }
                        #endregion Valores dependiendo del objeto
                    }
                    catch
                    { }
                    #region Matriz Detalle
                    switch (pVal.ItemUID)
                    {
                        #region Matrix Service              
                        case "mtServ":
                            switch (pVal.ColUID)
                            {
                                case "AcctCode":
                                    this.Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific).Value = val;
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("ActId").Cells.Item(pVal.Row).Specific).Value = AcctCode;
                                    
                                    break;
                                case "TaxCode":
                                    this.Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific).String = val;
                                    break;
                                case "OcrCode2":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode3":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode4":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode5":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "Project":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Project").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                            }
                            break;
                        #endregion Matrix Service

                        #region Matrix Item              
                        case "mtItem":
                            switch (pVal.ColUID)
                            {
                                case "ItemCode":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "TaxCode":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("TaxCode").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode").Cells.Item(pVal.Row).Specific).String = val;
                                    break;
                                case "OcrCode2":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode2").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode3":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode3").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode4":
                                    Matrix0.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode4").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "OcrCode5":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("OcrCode5").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                                case "Project":
                                    this.Matrix1.FlushToDataSource();
                                    ((SAPbouiCOM.EditText)this.Matrix1.Columns.Item("Project").Cells.Item(pVal.Row).Specific).Value = val;
                                    break;
                            }
                            break;
                            #endregion Matrix Item
                    }
                    #endregion Matriz Detalle
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("SelectChooise :" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private SAPbouiCOM.Matrix Matrix1;

        private void Matrix1_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SelectChooise(pVal, sboObject);
        }

        private void ComboBox1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string DocType = ((SAPbouiCOM.ComboBox)(this.UIAPIRawForm.Items.Item("DocType").Specific)).Value;
            switch (DocType)
            {
                case "I":
                    this.Matrix1.Item.Visible = true;
                    this.Matrix0.Item.Visible = false;
                    break;
                case "S":
                    this.Matrix0.Item.Visible = true;
                    this.Matrix1.Item.Visible = false;
                    break;
            }

        }

        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.LinkedButton LinkedButton1;
    }
}
