using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common.CRUD
{
    public class ASRDTE
    {
        public string DocumentID { get; set; }
        public string RutEmisor { get; set; }
        public string RznSoc { get; set; }
        public bool ExiEmisor { get; set; }
        public string TipoDTE { get; set; }
        public string Folio { get; set; }
        public DateTime FchEmis { get; set; }
        public DateTime FchVenc { get; set; }
        public string FmaPago { get; set; }
        public Double MntNeto { get; set; }
        public Double MntNto { get; set; }
        public Double MntExe { get; set; }
        public Double TasaIVA { get; set; }
        public Double IVA { get; set; }
        public Double MntTotal { get; set; }
        public string FolioRefOC { get; set; }
        public string FolioSAPOC { get; set; }
        public string FolioRefEM { get; set; }
        public string FolioSAPEM { get; set; }
        public string FolioRefFA { get; set; }
        public string FolioSAPFA { get; set; }
        public string DocEntryS { get; set; }
        public string ObjType { get; set; }
        public string XML { get; set; }
        public string PDF64 { get; set; }

    }
}
