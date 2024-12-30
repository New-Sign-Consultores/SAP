using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Net;

namespace DTERECEP
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;

                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                Conex.oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                AddStructAddon();
                Application.SBO_Application.StatusBar.SetText("Addon recepción de compras iniciado con exito.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oApp.Run();

                ///Nueva conexion
                #region Nueva conexion

                //SAPbouiCOM.Application oAppSBO;
                //string strConexion = ""; //variable que almacena el codigo de identificacion de conexion con SBO
                //string[] strArgumentos = new string[4];
                //SAPbouiCOM.SboGuiApi oSboGuiApi = null; //Variable para obtener la instacia activa de SBO

                //oSboGuiApi = new SAPbouiCOM.SboGuiApi();//Instancia nueva para la gestion de la conexion
                //strArgumentos = System.Environment.GetCommandLineArgs();//obtenemos el codigo de conexion del entorno configurado en "Propiedades -> Depurar -> Argumentos de la linea de comandos"

                //if (strArgumentos.Length > 0)
                //{
                //    if (strArgumentos.Length > 1)
                //    {
                //        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                //        if (strArgumentos[0].LastIndexOf("\\") > 0) strConexion = strArgumentos[1];
                //        else strConexion = strArgumentos[0];
                //    }
                //    else
                //    {
                //        //Verificamos que la aplicacion se este ejecutando en un ambiente SBO
                //        if (strArgumentos[0].LastIndexOf("\\") > -1) strConexion = strArgumentos[0];
                //        else Application.SBO_Application.MessageBox("debe tener SBO activo");
                //    }
                //}
                //else Application.SBO_Application.MessageBox("debe tener SBO activo");

                //oSboGuiApi.Connect(strConexion);//Establecemos la conexion
                //oAppSBO = oSboGuiApi.GetApplication(-1);//Asignamos la conexion a la aplicacion
                //Conex.oCompany = (SAPbobsCOM.Company)oAppSBO.Company.GetDICompany();


                //Menu MyMenuS = new Menu();
                //MyMenuS.AddMenuItems();
                ////Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(MyMenuS.SBO_Application_AppEvent);
                //Application.SBO_Application.MenuEvent += MyMenuS.SBO_Application_MenuEvent);
                ////oAppSBO. RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                //Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                //Conex.oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                //AddStructAddon();
                //Application.SBO_Application.StatusBar.SetText("Addon recepción de compras iniciado con exito.!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                ////oAppSBO. Run();


                #endregion Nueva conexion


            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

        static void AddStructAddon()
        {
            #region Security
            Common.Security sec = new Common.Security();
            sec.CrearTablaUsuario("SECURITY", "Security", SAPbobsCOM.BoUTBTableType.bott_NoObject, Application.SBO_Application);
            sec.crearCampo("SECURITY", "KEY", "KEY", SAPbobsCOM.BoFieldTypes.db_Memo, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", false, null);
            sec.crearCampo("SECURITY", "ValidTo", "Valid to", SAPbobsCOM.BoFieldTypes.db_Date, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", false, null);
            sec.RegistraAddon();
            #endregion Security

            Common.StructureLoad sl = new Common.StructureLoad();

            #region Add Tables
            sl.addUserTable("ASRDTE", "Documento DTE", SAPbobsCOM.BoUTBTableType.bott_MasterData, Application.SBO_Application);
            sl.addUserTable("ASCFRC", "Configuración recepción", SAPbobsCOM.BoUTBTableType.bott_MasterData, Application.SBO_Application);
            #endregion Add Tables

            #region Add Fields

            #region Fields - Configuración recepción
            sl.AddUserField("ASCFRC", "URLDTELIST", "Url Lista DTE", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "URLDTE", "Url Descarga DTE", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "URLACD", "Url Aceptación DTE", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "URLRZD", "Url Rechazo DTE", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "ENVIRONMENT", "Ambiente", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "RUTREC", "Rut receptor", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "USER", "Usuario", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "KEY", "Clave", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "SENTACD", "Enviar aceptación", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "SENTRZD", "Enviar rechazo", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "LKPRCOR", "Enlace para orden de compra", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASCFRC", "LKPRCDN", "Enlace para entrada de mercancia", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            #endregion Fields - Configuración recepción

            #region Fields - Documento DTE
            sl.AddUserField("ASRDTE", "DocumentID", "Document ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "RutEmisor", "Rut emisor", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "RznSoc", "Razon social", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "ExiEmisor", "Existe emisor", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "TipoDTE", "Tipo DTE", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "Folio", "Folio", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FchEmis", "Fecha Emisión", SAPbobsCOM.BoFieldTypes.db_Date, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FchVenc", "Fecha vencimiento", SAPbobsCOM.BoFieldTypes.db_Date, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FmaPago", "Forma pago", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "MntNeto", "Monto neto", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "MntExe", "Monto exento", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "TasaIVA", "Tasa IVA", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "IVA", "Monto IVA", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "MntTotal", "Monto total", SAPbobsCOM.BoFieldTypes.db_Float, 11, SAPbobsCOM.BoFldSubTypes.st_Sum, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioRefOC", "Folio Ref OC", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioSAPOC", "Folio SAP OC", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioRefEM", "Folio Ref EM", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioSAPEM", "Folio SAP EM", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioRefFA", "Folio Ref FA", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "FolioSAPFA", "Folio SAP FA", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "DocEntryS", "Id Interno", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "ObjType", "Tipo de documento", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "XML", "Document XML", SAPbobsCOM.BoFieldTypes.db_Memo, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            sl.AddUserField("ASRDTE", "PDF64", "Document PDF", SAPbobsCOM.BoFieldTypes.db_Memo, 0, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", false, null);
            #endregion Fields - Documento DTE

            #endregion Add Fields

            #region Udo

            #region ASRDTE
            Common.StructureLoad.ServiceUdo srvUddo;
            srvUddo.CanCancel = false;
            srvUddo.CanDelete = false;
            srvUddo.CanFind = true;
            srvUddo.CanLog = true;
            srvUddo.LogTableName = "AASRDTE";
            Common.StructureLoad.PrmtzionIU prmUd = new Common.StructureLoad.PrmtzionIU();
            prmUd.CanDefaultForm = false;
            prmUd.EnhancedForm = true;
            List<Common.StructureLoad.FindColumns> ltFindCol = new List<Common.StructureLoad.FindColumns>() { };
            List<Common.StructureLoad.FormDefault> ltFormDefa = new List<Common.StructureLoad.FormDefault>() { };
            //List<Common.StructureLoad.FormChild> LtFormChild = new List<Common.StructureLoad.FormChild>(){};
            List<Common.StructureLoad.ChildTables> ltchildTables = new List<Common.StructureLoad.ChildTables>() { };
            sl.AddUdo("ASRDTE", "Documentos", SAPbobsCOM.BoUDOObjType.boud_MasterData, "ASRDTE", srvUddo, prmUd, ltFindCol, ltFormDefa, ltchildTables, null);
            #endregion ASRDTE

            #region ASCFRC
            Common.StructureLoad.ServiceUdo srvUdoCF;
            srvUdoCF.CanCancel = false;
            srvUdoCF.CanDelete = false;
            srvUdoCF.CanFind = true;
            srvUdoCF.CanLog = true;
            srvUdoCF.LogTableName = "AASCFRC";
            Common.StructureLoad.PrmtzionIU prmUdCF = new Common.StructureLoad.PrmtzionIU();
            prmUdCF.CanDefaultForm = false;
            prmUdCF.EnhancedForm = true;
            List<Common.StructureLoad.FindColumns> ltFindColCF = new List<Common.StructureLoad.FindColumns>() { };
            List<Common.StructureLoad.FormDefault> ltFormDefaCF = new List<Common.StructureLoad.FormDefault>() { };
            //List<Common.StructureLoad.FormChild> LtFormChildCF = new List<Common.StructureLoad.FormChild>() { };
            List<Common.StructureLoad.ChildTables> ltchildTablesCF = new List<Common.StructureLoad.ChildTables>() { };
            sl.AddUdo("ASCFRC", "Configuración recepción", SAPbobsCOM.BoUDOObjType.boud_MasterData, "ASCFRC", srvUdoCF, prmUdCF, ltFindColCF, ltFormDefaCF, ltchildTablesCF, null);
            #endregion ASCFRC

            #endregion Udo

            #region Set Udo

            #region ASCFRC
            Common.StructureLoad.UDo uDoCFRC = new Common.StructureLoad.UDo();
            uDoCFRC.Code = "REF01";
            uDoCFRC.NameUDO = "ASCFRC";
            uDoCFRC.Table = "ASCFRC";
            List<List<string>> FieldCFRC = new List<List<string>>
            {
                new List<string>{"Code", "REF01" },
                new List<string>{"Name", "Configuración 1" }
            };
            uDoCFRC.FieldValue = FieldCFRC;
            sl.SetValorUdo(uDoCFRC);
            #endregion ASCFRC

            #endregion Set Udo

        }

        /// <summary>
        /// Agrega una entrada a un Menu de SBO
        /// <alert class="note">
        /// <para>Para mayor sobre los tipo de menu referencia de los tipos revisar la ayuda del SDK de SBO</para>
        /// </alert>
        /// </summary>
        /// <param name="UniqueId">Identificador Unico del Menu que será creado</param>
        /// <param name="Name">Nombre que sera mostrado en SBO</param>
        /// <param name="PrincipalMenuId">Identificador Unico del Menu que contendra la nueva entrada</param>
        /// <param name="type">Tipo de Menu</param>
        private void CreaMenu(string uniqueId, string name, string principalMenuId, SAPbouiCOM.BoMenuType type)
        {
            SAPbouiCOM.MenuCreationParams objParams;
            SAPbouiCOM.Menus objSubMenu;
            int posmenu = 0;
            try
            {
                objSubMenu = Application.SBO_Application.Menus.Item(principalMenuId).SubMenus;

                if (Application.SBO_Application.Menus.Exists(uniqueId) == false)
                {
                    posmenu = objSubMenu.Count;
                    objParams = (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    objParams.Type = type;
                    objParams.UniqueID = uniqueId;
                    objParams.String = name;
                    objParams.Position = posmenu + 1;                    
                    objSubMenu.AddEx(objParams);
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox( "Add Menu " + ex.Message);
            }
        }
    }
}
