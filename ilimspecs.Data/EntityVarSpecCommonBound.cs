using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //Класс в который записываются спецификации которые имеют общие границы
    public class EntityVarSpecCommonBound
    {
        public string Var_Id { get; set; }

        public decimal? L_Reject { get; set; }
        public decimal? L_Warning { get; set; }
        public decimal? Target { get; set; }
        public decimal? U_Reject { get; set; }
        public decimal? U_Warning { get; set; }
        public int? CountProduct { get; set; }

        public decimal? LoLo { get { return L_Reject; } }
        public decimal? Lo { get { return L_Warning; } }
        public decimal? HiHi { get { return U_Reject; } }
        public decimal? Hi { get { return U_Warning; } }
    }
}
