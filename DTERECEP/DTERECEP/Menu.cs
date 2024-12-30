using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace DTERECEP
{
    class Menu
    {
        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "DTERECEP";
            oCreationPackage.String = "Recepcion Compras";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch 
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("DTERECEP");
                oMenus = oMenuItem.SubMenus;
                try
                {
                    // Create s sub menu
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "FShppingRcpt";
                    oCreationPackage.String = "Recepción de compras";
                    oMenus.AddEx(oCreationPackage);
                }
                catch { }
                try
                {
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "FormLicAddon";
                    oCreationPackage.String = "Licencia Addon";
                    oMenus.AddEx(oCreationPackage);
                }
                catch { }
                try
                {
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "FCASCFRC";
                    oCreationPackage.String = "Configuración recepción";
                    oMenus.AddEx(oCreationPackage);
                }
                catch { }
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists" + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                DTERECEP.Common.Security sec = new DTERECEP.Common.Security();
                if (pVal.BeforeAction )
                {
                    switch (pVal.MenuUID)
                    {
                        case "FShppingRcpt":
                            //if (sec.ValidLic()) //Licencia Valida
                            {
                                Forms.FShppingRcpt fShppingRcpt = new Forms.FShppingRcpt();
                                fShppingRcpt.Show();
                            }
                            //else Application.SBO_Application.MessageBox("Debe asignar una licencia valida.\n Favor comuniquese con su proveedor.!");
                            break;
                        case "FCASCFRC":
                            //if (sec.ValidLic()) //Licencia Valida
                            {
                                Forms.FCASCFRC fCASCFRC = new Forms.FCASCFRC();
                                fCASCFRC.Show();
                            }
                            //else Application.SBO_Application.MessageBox("Debe asignar una licencia valida.\n Favor comuniquese con su proveedor.!");
                            break;
                        case "FormLicAddon":
                            Forms.FormLicAddon formLicAddon = new Forms.FormLicAddon();
                            formLicAddon.Show();
                            break;
                    }
                }

                //Common.Security sec = new Common.Security();
                //if (sec.ValidLic()) //Licencia Valida
                //{
                //    Forms.FormChangeFol activeForm = new Forms.FormChangeFol();
                //    activeForm.Show();
                //}
                //else
                //{
                //    Application.SBO_Application.MessageBox("Debe asignar una licencia valida.\n Favor comuniquese con su proveedor.!");
                //}
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

    }
}
