using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

namespace DTERECEP.Common
{
    public class Security
    {
        public bool CrearTablaUsuario(string Nombre, string descripcion, SAPbobsCOM.BoUTBTableType tipo, SAPbouiCOM.Application oApli)
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
                try
                {
                    Marshal.ReleaseComObject(oUTables);
                }
                catch
                {
                    GC.Collect();
                }
            }
            return existe;
        }
        public void crearCampo(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo, int tamaño, SAPbobsCOM.BoFldSubTypes subtipo, string ValorPorDefecto, string sLinkedTable, Boolean Mandatory, List<List<String>> ValidValues)//CHG
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery("select COUNT(1) from CUFD where (tableID='" + tabla + "' or tableID='@" + tabla + "') and AliasID='" + campo + "'");

            existeCampo = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;

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

                    if (sLinkedTable.ToString() != "")
                        oCampo.LinkedTable = sLinkedTable;

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
        public bool RegistraAddon()
        {
            bool exist = false;
            string AddonName = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
            SAPbobsCOM.Recordset oRec;
            try
            {
                string query = @"SELECT T0.""U_KEY"", T0.""U_ValidTo"", T0.""Code"" FROM ""@SECURITY""  T0 WHERE T0.""Code"" ='" + AddonName + "' ";
                oRec = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(query);
                if (oRec.RecordCount.Equals(0))
                {
                    query = @"
                                INSERT INTO ""@SECURITY""
                                           (""Code""
                                           ,""Name""
                                           ,""U_KEY""
                                           ,""U_ValidTo"")
                                     VALUES
                                           ( '" + AddonName + @"'
                                           , '" + AddonName + @"'
                                           , 'XXXXXXXXXXXXXXXXXXXXXXXXXXXX' 
                                           , '1900-01-01'   )
                               ";
                    oRec.DoQuery(query);
                }

            }
            catch//(Exception ex)
            {
                exist = false;
                //oApli.MessageBox("ERROR creartabla : " + er.Message + " Source " + er.Source + " Stack " + er.StackTrace + " " + er.TargetSite, 1, "", "", "");
            }
            return exist;
        }
        public bool ValidLic()
        {
            bool valid = false;
            try
            {
                string AddonName = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
                string KEY = null, Server = null, NroInsT = null;
                string Query = @"SELECT ""Code""
                                      ,""U_KEY""
                                  FROM ""@SECURITY""
                                WHERE ""Code"" IN ('" + AddonName + @"')
                                ";
                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(Query);
                if (oRec.RecordCount > 0)
                {
                    KEY = Convert.ToString(oRec.Fields.Item("U_KEY").Value);
                    Server = Conex.oCompany.Server;
                    NroInsT = SAPbouiCOM.Framework.Application.SBO_Application.Company.InstallationId;
                    //var  sys = SAPbouiCOM.Framework.Application.SBO_Application.Company.SystemId;
                    ValidLic.Lic Lic = new ValidLic.Lic();
                    valid = Lic.ValidLic(KEY, AddonName, Server, NroInsT);
                }
                else
                {
                    valid = false;
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Error al validar licencia.!" + ex.Message);
                valid = false;
            }
            return valid;
        }
    }
}