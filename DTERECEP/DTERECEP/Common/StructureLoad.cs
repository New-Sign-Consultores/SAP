using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common
{
    public class StructureLoad
    {
        public bool addUserTable(string Nombre, string descripcion, SAPbobsCOM.BoUTBTableType tipo, SAPbouiCOM.Application oApli)
        {
            bool existe = false;
            SAPbobsCOM.UserTablesMD oUTables = null;
            int error = 0;
            try
            {
                GC.Collect();
                oUTables = (SAPbobsCOM.UserTablesMD)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUTables.GetByKey(Nombre))
                {
                    oUTables = null;
                    existe = true;
                }
                else
                {
                    oUTables.TableName = Nombre;
                    oUTables.TableDescription = descripcion;
                    oUTables.TableType = tipo;
                    error = oUTables.Add();
                    if (error != 0)
                    {
                        oApli.MessageBox("ERROR company : " + Conex.oCompany.GetLastErrorDescription(), 1, "", "", "");
                    }
                }
            }
            catch (Exception er)
            {
                oApli.MessageBox("ERROR creartabla : " + er.Message + " Source " + er.Source + " Stack " + er.StackTrace + " " + er.TargetSite, 1, "", "", "");
            }
            finally
            {
                //try
                //{
                //    //Marshal.ReleaseComObject(oUTables);
                //}
                //catch
                //{
                GC.Collect();
                //}
            }
            return existe;
        }

        public void AddUserField(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo, int tamaño, SAPbobsCOM.BoFldSubTypes subtipo, string ValorPorDefecto, string sLinkedTable, string LinkedUDO, Boolean Mandatory, List<List<String>> ValidValues)//CHG
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT COUNT(1) FROM CUFD WHERE ""TableID""='@" + tabla + @"' AND ""AliasID"" ='" + campo + @"' ";
            if (tabla.Length.Equals(4))
                query = @"SELECT COUNT(1) FROM CUFD WHERE ""TableID""='" + tabla + @"' AND ""AliasID"" ='" + campo + @"' ";
            else
                query = @"SELECT COUNT(1) FROM CUFD WHERE ""TableID""='@" + tabla + @"' AND ""AliasID"" ='" + campo + @"' ";
            rs.DoQuery(query);

            existeCampo = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();
            SAPbobsCOM.UserFieldsMD oCampo = (SAPbobsCOM.UserFieldsMD)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if ((existeCampo == 0))
                {
                    oCampo.TableName = tabla;
                    oCampo.Name = campo;
                    oCampo.Description = descripcion;
                    oCampo.Type = tipo;
                    oCampo.SubType = subtipo;
                    oCampo.Mandatory = Mandatory ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    if (tamaño > 0)
                    {
                        oCampo.EditSize = tamaño;
                    }

                    //Linl Table
                    if (sLinkedTable.ToString() != "")
                        oCampo.LinkedTable = sLinkedTable;

                    //Link UDo
                    if (!string.IsNullOrEmpty(LinkedUDO))
                        oCampo.LinkedUDO = LinkedUDO;

                    if (ValidValues != null)
                    {
                        foreach (List<String> ValidValue in ValidValues)
                        {
                            oCampo.ValidValues.Value = ValidValue[0];
                            oCampo.ValidValues.Description = ValidValue[1];
                            oCampo.ValidValues.Add();
                        }
                    }

                    if (ValorPorDefecto.ToString() != "")
                    {
                        oCampo.DefaultValue = ValorPorDefecto;
                    }

                    int RetVal = oCampo.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Conex.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCampo);
            }
        }

        public void AddUdo(string Code, string Name, SAPbobsCOM.BoUDOObjType ObjectType, string TableName, ServiceUdo servUdo, PrmtzionIU prmtzionIU, List<FindColumns> LtFindColumns, List<FormDefault> LtFormDefault, List<ChildTables> ltchildTables, string FormSRF)
        {
            int existeudo = 0;
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(@"SELECT COUNT(1) FROM OUDO WHERE ""Code"" = '" + Code + "' ");

            existeudo = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;
            GC.Collect();
            SAPbobsCOM.IUserObjectsMD oUOMDo = (SAPbobsCOM.IUserObjectsMD)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            try
            {
                if ((existeudo == 0))
                {
                    oUOMDo.Code = Code;
                    oUOMDo.Name = Name;
                    oUOMDo.ObjectType = ObjectType;
                    oUOMDo.TableName = TableName;

                    ///Modificación servicios
                    oUOMDo.CanFind = servUdo.CanFind ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUOMDo.CanDelete = servUdo.CanDelete ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUOMDo.CanCancel = servUdo.CanCancel ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUOMDo.CanLog = servUdo.CanLog ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    if (!string.IsNullOrEmpty(servUdo.LogTableName))
                        oUOMDo.LogTableName = servUdo.LogTableName;

                    if (!string.IsNullOrEmpty(FormSRF))
                        oUOMDo.FormSRF = FormSRF;

                    ///Parametrizaciones de IU
                    if (prmtzionIU.EnhancedForm)
                    {
                        oUOMDo.CanCreateDefaultForm = prmtzionIU.CanDefaultForm ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                        oUOMDo.EnableEnhancedForm = prmtzionIU.EnhancedForm ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                        oUOMDo.MenuItem = prmtzionIU.MenuItem ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                        oUOMDo.MenuCaption = prmtzionIU.MenuCaption;
                        oUOMDo.FatherMenuID = prmtzionIU.FatherMenuID;
                        oUOMDo.Position = prmtzionIU.position;
                        oUOMDo.MenuUID = Code;
                    }

                    ///Modificando los campos para "Buscar"
                    foreach (FindColumns findCol in LtFindColumns)
                    {
                        oUOMDo.FindColumns.ColumnAlias = findCol.ColumnAlias;
                        oUOMDo.FindColumns.ColumnDescription = findCol.ColumnDescription;
                        oUOMDo.FindColumns.Add();
                    }

                    ///Modificando los campos para el formulario por defecto
                    foreach (FormDefault formDefault in LtFormDefault)
                    {
                        oUOMDo.FormColumns.Editable = formDefault.Editable ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                        oUOMDo.FormColumns.FormColumnAlias = formDefault.FColAlias;
                        oUOMDo.FormColumns.FormColumnDescription = formDefault.FColDescription;
                        oUOMDo.FormColumns.Add();
                    }

                    int child = 0;
                    ///Enlazando tablas de usuarios subordinadas ChildTables
                    foreach (ChildTables childTables in ltchildTables)
                    {
                        oUOMDo.ChildTables.TableName = childTables.ChildTableName;
                        oUOMDo.ChildTables.LogTableName = childTables.LogTableName;
                        oUOMDo.ChildTables.Add();
                        ++child;

                        ///Fijando campos para un formulario por defecto de tabla subordinada
                        foreach (FormChild formChild in childTables.ltFormChild)
                        {
                            oUOMDo.EnhancedFormColumns.ColumnAlias = formChild.FColAlias;
                            oUOMDo.EnhancedFormColumns.ColumnDescription = formChild.FColDescription;
                            oUOMDo.EnhancedFormColumns.Editable = formChild.Editable ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                            oUOMDo.EnhancedFormColumns.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES;
                            oUOMDo.EnhancedFormColumns.ChildNumber = child;
                            oUOMDo.EnhancedFormColumns.Add();

                        }
                    }

                    //SAPbobsCOM~ChildTables_(UserObjectMD_ChildTables)
                    //oUOMDo.FormColumns.SonNumber

                    int RetVal = oUOMDo.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Conex.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUOMDo);
            }
        }

        public void SetValorUdo(UDo UdoSet)
        {
            #region ConfiAddon
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //select a la tabla para saber si existe ya la config
            string query = @" SELECT T0.""Code"" FROM ""@" + UdoSet.Table + @"""  T0 WHERE T0.""Code"" = '" + UdoSet.Code + @"' ";
            rs.DoQuery(query);

            string Code = Convert.ToString(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;
            if (string.IsNullOrEmpty(Code))
            {
                #region Variables
                SAPbobsCOM.GeneralService oDocGeneralService = null;
                SAPbobsCOM.CompanyService oCompService = null;
                SAPbobsCOM.GeneralData oDocGeneralData = null;
                SAPbobsCOM.GeneralDataCollection oGeneralDataCollection = null;
                SAPbobsCOM.GeneralData oSon = null;

                #endregion Variables
                try
                {
                    try
                    {
                        #region Head
                        Conex.oCompany.StartTransaction();
                        oCompService = Conex.oCompany.GetCompanyService();
                        // toma los parámetros de la udo
                        oDocGeneralService = oCompService.GetGeneralService(UdoSet.NameUDO);
                        oDocGeneralData = (SAPbobsCOM.GeneralData)oDocGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                        foreach (List<String> fieldValue in UdoSet.FieldValue)
                        {
                            oDocGeneralData.SetProperty(fieldValue[0], fieldValue[1]);
                        }
                        //oDocGeneralData.SetProperty("U_AcepE", "Y");

                        #endregion Head

                        #region Collection
                        if (UdoSet.LChild != null)
                        {
                            foreach (LTChild ltChild in UdoSet.LChild)
                            {
                                oGeneralDataCollection = oDocGeneralData.Child(ltChild.NameChild);
                                oSon = oGeneralDataCollection.Add();
                                foreach (List<String> Child in ltChild.FieldValue)
                                {
                                    oSon.SetProperty(Child[0], Child[1]);
                                }
                            }
                        }
                        #endregion Collection

                        oDocGeneralService.Add(oDocGeneralData);
                        if (Conex.oCompany.InTransaction)
                        {
                            Conex.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Configuración creada exitosamente.!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompService);
                        oCompService = null;
                        GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        string Error = ex.Message;
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Error al crear configuración :" + Error + ".!", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        if (Conex.oCompany.InTransaction)
                        {
                            Conex.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                    }
                }
                finally
                {
                    GC.Collect();
                }
            }
            #endregion ConfiAddon
        }
        public struct ServiceUdo
        {
            public bool CanFind;
            public bool CanDelete;
            public bool CanCancel;
            public bool CanLog;
            public string LogTableName;
        }

        /// <summary>
        /// Parametrizaciones de IU
        /// </summary>
        public struct PrmtzionIU
        {
            /// <summary>
            /// Formulario por defecto
            /// </summary>
            public bool CanDefaultForm;

            /// <summary>
            /// Estilo de línea de cabecera
            /// </summary>
            public bool EnhancedForm;

            /// <summary>
            /// /Opción de menú
            /// </summary>
            public bool MenuItem;

            /// <summary>
            /// Título de menú
            /// </summary>
            public string MenuCaption;

            /// <summary>
            /// ID de menú superior
            /// </summary>
            public int FatherMenuID;

            /// <summary>
            /// position
            /// </summary>
            public int position;
        }

        /// <summary>
        /// Modificando los campos para "Buscar"
        /// </summary>
        public struct FindColumns
        {
            /// <summary>
            /// Alias columna
            /// </summary>
            public string ColumnAlias;

            /// <summary>
            /// Descripción columna
            /// </summary>
            public string ColumnDescription;
        }

        /// <summary>
        /// Formulario por defecto
        /// </summary>
        public struct FormDefault
        {
            /// <summary>
            /// Columna editable
            /// </summary>
            public bool Editable;

            /// <summary>
            /// Alias columna
            /// </summary>
            public string FColAlias;

            /// <summary>
            /// Descripción columna
            /// </summary>
            public string FColDescription;
        }

        /// <summary>
        /// Enlazando tablas de usuarios subordinadas
        /// </summary>
        public struct ChildTables
        {
            /// <summary>
            /// Nombre de la tabla subordinada
            /// </summary>
            public string ChildTableName;

            /// <summary>
            /// Nombre de la tabla log subordinada
            /// </summary>
            public string LogTableName;

            public List<FormChild> ltFormChild;

        }

        /// <summary>
        /// Formulario por defecto tablas subordinadas
        /// </summary>
        public struct FormChild
        {
            /// <summary>
            /// Columna editable
            /// </summary>
            public bool Editable;

            /// <summary>
            /// Alias columna
            /// </summary>
            public string FColAlias;

            /// <summary>
            /// Descripción columna
            /// </summary>
            public string FColDescription;
        }

        public struct UDo
        {
            public string NameUDO;
            public string Code;
            public string Table;
            public List<List<String>> FieldValue;
            public List<LTChild> LChild;

        }

        public struct LTChild
        {
            public string NameChild;
            public List<List<String>> FieldValue;
        }
    }
}
