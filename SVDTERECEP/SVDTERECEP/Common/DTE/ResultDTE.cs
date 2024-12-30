using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTERECEP.Common.DTE
{
    public class ResultDTE
    {
        public int Id { get; set; }
        public string Mensaje { get; set; }       
        public bool Success { get; set; }
        //public DataSet ds;
        public DTERECEP.DTE.DTE DTE { get; set; }
        public string XMLString;
        
    }
}
