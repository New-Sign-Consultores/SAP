using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common.NEWSIGN
{
    class SearchResultXML
    {
        public Data Data { get; set; }
        public string Description { get; set; }
        public string Result { get; set; }
        public string StackTrace { get; set; }
        public string SearchTime { get; set; }
        public int TotalDocuments { get; set; }

        public SearchResultXML()
        {

        }
    }

    class Data
    {
        public string stringData { get; set; }
        public List<DocumentXML> ltDocuments { get; set; }
        public Data()
        {
            ltDocuments = new List<DocumentXML>();
        }
    }
    public class DocumentXML
    {
        public string RecipientRUT { get; set; }
        public string IssuerName { get; set; }
        public DateTime Created { get; set; }
        public string IssuerRUT { get; set; }
        public string ExternalID { get; set; }
        public bool Replaced { get; set; }
        public string GetUniqueBusinessId { get; set; }
        public string DocumentReferences { get; set; }
        public string Statuses { get; set; }
        public string Version { get; set; }
        public double TotalAmount { get; set; }
        public bool Deleted { get; set; }
        public DateTime EmissionDate { get; set; }
        public string MailReferences { get; set; }
        public string DTEType { get; set; }
        public double Iva { get; set; }
        public string Folio { get; set; }
        public string RecipientName { get; set; }
        public DateTime Modified { get; set; }
        public string DocumentID { get; set; }
        public string Id { get; set; }
        public string FchRespComercial { get; set; }
        public double NetoAmount { get; set; }
        public string CanalRecepcion { get; set; }
        public string Anulado { get; set; }
        public string Eliminado { get; set; }
        public DateTime FchRecepSII { get; set; }
        public string Firma { get; set; }
        public string Desanulado { get; set; }
        public string Intercambio { get; set; }
        public string AutorizadoSII { get; set; }
        public string Recibido { get; set; }
        public string Distribuido { get; set; }
        public string CmnaRecep { get; set; }
        public bool TieneArchivo { get; set; }
        public string AnuladoContable { get; set; }
        public string RUTRecep { get; set; }
        public string Grupo { get; set; }
        public string Elaboracion { get; set; }
        public double MntNeto { get; set; }
        public string ID { get; set; }
        public string CiudadOrigen { get; set; }
        public double MntTotal { get; set; }
        public string Estructura { get; set; }
        public int FmaPago { get; set; }
        public DateTime FchEmis { get; set; }
        public string ErrorPrint { get; set; }
        public string CmnaOrigen { get; set; }
        public string CdgSIISucur { get; set; }
        public string Contacto { get; set; }
        public string CiudadRecep { get; set; }
        public string RUTEmisor { get; set; }
        public string CEN { get; set; }
        public string Aprobado { get; set; }
        public string Cesion { get; set; }
        public DateTime TimeStamp { get; set; }
        public string TipoDocumento { get; set; }
        public string AprobadoSII { get; set; }
        public double IVA { get; set; }
        public string RznSoc { get; set; }
        public string TipoDTE { get; set; }
        public string NmbItem { get; set; }
        public string Conciliado { get; set; }
        public string Procesado { get; set; } 
        public string RznSocRecep { get; set; }
        public string DscItem { get; set; }
        public string DownloadCustomerDocumentUrl { get; set; }
        public string ClaimAction { get; set; }
        public DocumentXML()
        {

        }
    }
}
