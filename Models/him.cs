using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Inventariz.Models
{
    public class him
    {
        public Int32 id { get; set; }
        public System.DateTime dates { get; set; }
        public Nullable<double> P { get; set; }
        public Nullable<double> temp { get; set; }
        public Nullable<double> P20 { get; set; }
        public Nullable<double> water { get; set; }
        public Nullable<double> saltmg { get; set; }
        public Nullable<double> masend { get; set; }
        public Nullable<double> mechan { get; set; }
    }
}