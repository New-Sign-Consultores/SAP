using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DTERECEP.Forms
{
    [FormAttribute("DTERECEP.Forms.FormLicAddon", "Forms/FormLicAddon.b1f")]
    class FormLicAddon : UserFormBase
    {
        public FormLicAddon()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btSearch").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tbPathFile").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("btLoad").Specific));
            this.Button1.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button1_ClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        private void OnCustomInitialize()
        {

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            DTERECEP.Common.FileManager fm = new Common.FileManager();
            fm.OpenFile(oForm, "tbPathFile");
        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button1;

        private void Button1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(this.UIAPIRawForm.UniqueID);
            Common.FileManager fm = new Common.FileManager();
            fm.UploadLicense(this.EditText0.Value);

        }
    }
}
