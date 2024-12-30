using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace DTERECEP.Forms
{
    [FormAttribute("DTERECEP.Forms.FCASCFRC", "Forms/FCASCFRC.b1f")]
    class FCASCFRC : UserFormBase
    {
        public FCASCFRC()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("RUT").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("ENV").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Code").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("SENTACD").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("SENTRZD").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_17").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("UDTELIST").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("URLDTE").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("URLACD").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("URLRZD").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("USER").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("KEY").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("LKPRCOR").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("LKPRCDN").Specific));
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
            this.EditText2.Item.Visible = false;
            this.ComboBox0.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            this.UIAPIRawForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
            this.EditText1.Value = "*";
            this.UIAPIRawForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.EditText EditText9;
    }
}
