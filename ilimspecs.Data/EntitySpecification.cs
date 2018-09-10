using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //class EntitySpecification
    //{
    //}
    public class EntitySpecification
    {
        public string Name { get; set; }
        public string ProductionUnit { get; set; }
        public string PathAF { get; set; }
        public string Tag { get; set; } // Имя тега который хранится в атрибуте "Тэг PI"

        public List<ProductLimits> ProductLimits = new List<ProductLimits>();
    }

    public class ProductLimits
    {
        public Int64 VS_Id { get; set; }
        public DateTime Date_Effective { get; set; }
        public DateTime? Date_Expiration { get; set; }
        public decimal? L_Reject { get; set; }
        public decimal? L_Warning { get; set; }
        public Int64 Product_Id { get; set; }
        public decimal? Target { get; set; }
        public decimal? U_Reject { get; set; }
        public decimal? U_Warning { get; set; }
        public string Var_Id { get; set; }
        public string Var_Type { get; set; }
        public string Author { get; set; }

        public decimal? LoLo { get { return L_Reject; } }
        public decimal? Lo { get { return L_Warning; } }
        public decimal? HiHi { get { return U_Reject; } }
        public decimal? Hi { get { return U_Warning; } }

        //public Int64 Product_Id { get; set; }
        public string Extended_Info { get; set; }
        public bool Product_Is_Active { get; set; }
        public string Product_Code { get; set; }
        public string Product_Desc { get; set; }
        public int ProductionUnit_Id { get; set; }
        public string Product_SAP { get; set; }
        public string Product_Procont { get; set; }
    }
}
