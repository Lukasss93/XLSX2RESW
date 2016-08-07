using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSX2RESW
{
    public class JProject
    {
        public string code { get; set; }
        public List<JValues> values { get; set; }
    }

    public class JValues
    {
        public string id { get; set; }
        public string value { get; set; }
    }
}
