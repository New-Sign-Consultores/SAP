using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SVDTERECEP
{
    class Conex
    {
        public static SAPbobsCOM.Company oCompany; 
        
    }
    public static class DTECompany
    {
        public static string ENVIRONMENT { get; set; }
        public static string RUTREC { get; set; }
        public static bool SENTACD { get; set; }
        public static bool SENTRZD { get; set; }
        public static string URLDTELIST { get; set; }
        public static string URLDTE { get; set; }
        public static string URLACD { get; set; }
        public static string URLRZD { get; set; }
        public static string USER { get; set; }
        public static string KEY { get; set; }
    }
}
