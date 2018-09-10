using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //Класс для записи границ по определенному продукту и определенному тегу
    public class ProductLimitCommondBound : ProductLimit
    {
        public Int64 VS_Id { get; set; }
        public Int64 Product_Id { get; set; }

        public ProductLimitCommondBound(string PrUnit_Desc, string Tg, string PrCode, string auth) : base(PrUnit_Desc, Tg, PrCode, auth)
        {
            
        }
    }
}
