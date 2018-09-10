using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    public class EntityProduct
    {
        public Int64 Product_Id { get; set; }
        public string Extended_Info { get; set; }
        public bool Product_Is_Active { get; set; }
        public string Product_Code { get; set; }
        public string Product_Desc { get; set; }
        public int ProductionUnit_Id { get; set; }
        public string Product_SAP { get; set; }
        public string ProductionUnit_Desc { get; set; }
        public string Product_Procont { get; set; }

        public override string ToString()
        {
            return Product_Desc;
        }

    }
}
