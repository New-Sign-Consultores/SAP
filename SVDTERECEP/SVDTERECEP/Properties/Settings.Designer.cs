﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Este código fue generado por una herramienta.
//     Versión de runtime:4.0.30319.42000
//
//     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
//     se vuelve a generar el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace SVDTERECEP.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.10.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute(@"SELECT 
 T0.""U_RutEmisor""
,T0.""U_TipoDTE""
,T0.""U_Folio""
,'N' AS ""Check""
,T0.""U_RznSoc""
,(CASE ISNULL(T1.""CardCode"",'N') WHEN 'N' THEN 'N' ELSE 'Y' END) AS ""U_ExiEmisor""
,T0.""U_FchEmis""
,T0.""U_FchVenc""
,T0.""U_FmaPago""
,T0.""U_MntNeto""
,T0.""U_MntExe""
,T0.""U_TasaIVA""
,T0.""U_IVA""
,T0.""U_MntTotal""
,CAST('' AS VARCHAR(254)) ""Glosa""
,T0.""U_FolioRefOC""
,T0.""U_FolioSAPOC""
,CAST(0 As float) AS ""MTotalOC""
,T0.""U_FolioRefEM""
,(CASE WHEN LEN(T0.""U_FolioSAPEM"") > 0 THEN T0.""U_FolioSAPEM"" ELSE NULL END)  AS ""U_FolioSAPEM""
,CAST(0 As float) AS ""MTotalEM""
,T0.""U_FolioRefFA""
,T0.""U_FolioSAPFA""
,0 AS ""CodRefNC""
,(CASE WHEN LEN(CAST(T0.""U_XML"" AS VARCHAR(MAX))) > 0  THEN '' ELSE 'Sin XML' END) AS ""Mensaje""
,T0.""U_DocEntryS""
,T0.""U_ObjType""
,T0.""U_XML""
,T0.""U_PDF64""
,T0.""Code""     
,T1.""CardCode""
,T0.""U_DocumentID""  


FROM ""@ASRDTE"" T0
LEFT JOIN ""OCRD"" T1 ON  CAST(T1.""LicTradNum"" AS varchar(20)) = CAST(T0.""U_RutEmisor"" AS varchar(20)) AND T1.""CardType"" ='S'
WHERE T0.""U_FchEmis"" BETWEEN FchDesd AND FchHast
AND T0.""U_TipoDTE"" IN (Documentos)")]
        public string ListaDTESQL {
            get {
                return ((string)(this["ListaDTESQL"]));
            }
            set {
                this["ListaDTESQL"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("MS2019")]
        public string DbServerType {
            get {
                return ((string)(this["DbServerType"]));
            }
            set {
                this["DbServerType"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("SBODemoCL")]
        public string CompanyDB {
            get {
                return ((string)(this["CompanyDB"]));
            }
            set {
                this["CompanyDB"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("DESKTOP-S39KKQH")]
        public string Server {
            get {
                return ((string)(this["Server"]));
            }
            set {
                this["Server"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("manager")]
        public string UserName {
            get {
                return ((string)(this["UserName"]));
            }
            set {
                this["UserName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("B1Admin*")]
        public string Password {
            get {
                return ((string)(this["Password"]));
            }
            set {
                this["Password"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string FDESDE {
            get {
                return ((string)(this["FDESDE"]));
            }
            set {
                this["FDESDE"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("")]
        public string FHASTA {
            get {
                return ((string)(this["FHASTA"]));
            }
            set {
                this["FHASTA"] = value;
            }
        }
    }
}
