using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DTERECEP.Common
{
    public class FileManager
    {
        private string pathItem;
        public SAPbouiCOM.Form DialogForm;
        public string filename;
        public SAPbobsCOM.UserFieldsMD oUserFieldsMD;

        public void OpenFile(SAPbouiCOM.Form oForm, string path)
        {
            try
            {
                pathItem = path;
                DialogForm = oForm;
                System.Threading.Thread ShowFolderBrowserThread;
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                //ShowFolderBrowserThread.Abort();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message, 1);
            }
        }

        private void ShowFolderBrowser()
        {
            try
            {
                NativeWindow nws = new NativeWindow();
                OpenFileDialog MyTest = new OpenFileDialog();
                MyTest.Multiselect = false;
                MyTest.Filter = "Text Files (.txt)|*.txt";
                Process[] MyProcs = null;
                //string filename = null;
                MyProcs = Process.GetProcessesByName("SAP Business One");
                nws.AssignHandle(System.Diagnostics.Process.GetProcessesByName("SAP Business One")[0].MainWindowHandle);
                if (MyTest.ShowDialog(nws) == System.Windows.Forms.DialogResult.OK)
                {
                    filename = MyTest.FileName;
                    SAPbouiCOM.EditText Texto1 = (SAPbouiCOM.EditText)DialogForm.Items.Item(pathItem).Specific;
                    Texto1.Value = filename;
                    DialogForm = null;
                    pathItem = null;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        public void UploadLicense(string Path)
        {
            string line;
            using (StreamReader file = new StreamReader(Path))
            {
                string addonName = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyTitleAttribute>().Title;
                line = file.ReadLine();
                ValidLic.Lic Lic = new ValidLic.Lic();
                DateTime dtVigencia = Lic.GetValidTo(line);
                string NameAdd = Lic.GetAddonName(line);
                if (NameAdd == addonName)
                {
                    SAPbobsCOM.Recordset rs0 = null;
                    rs0 = ((SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    string Q0 = @"UPDATE ""@SECURITY"" SET U_KEY = '" + line + @"', ""U_ValidTo"" = '" + dtVigencia.ToString("yyyy-MM-dd") + @"' WHERE ""Name"" = '" + addonName + @"'";
                    rs0.DoQuery(Q0);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs0);
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("Se cargó la licencia exitosamente.\nSe va a cerrar SAP Business One, por favor ingresar nuevamente a la aplicación");
                    SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("526");
                }
                else
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("La licencia seleccionada es inválida");
                }
                file.Close();
            }
        }

        private bool FieldExists(string sTableID, string sAliasID)
        {
            try
            {
                int FieldId = GetFieldID(sTableID, sAliasID);
                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                if (oUserFieldsMD.GetByKey(sTableID, FieldId) == true)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    GC.Collect();
                    return true;
                }
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    GC.Collect();
                    return false;
                }
            }
            catch //(Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                GC.Collect();
                return false;
            }
        }

        private int GetFieldID(string sTableID, string sAliasID)
        {
            int iRetVal = -1;
            try
            {
                SAPbobsCOM.Recordset rs0 = null;
                rs0 = ((SAPbobsCOM.Recordset)(Conex.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                string Q0 = ("select \"FieldID\" from CUFD where \"TableID\" = '" + sTableID + "' and \"AliasID\" = '" + sAliasID + "'");
                rs0.DoQuery(Q0);
                if (!rs0.EoF) iRetVal = Convert.ToInt32(rs0.Fields.Item("FieldID").Value.ToString());
                rs0 = null;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {

            }
            return iRetVal;
        }
    }
}
