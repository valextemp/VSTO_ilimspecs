using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.BL
{
    //Класс для учета продуктов на листе спецификаций
    public class ProductOnSheet
    {
        readonly string product_Code;
        readonly int columnProduct_Code;
        public string Product_Code { get { return product_Code; } }
        public int ColumnProduct_Code { get { return columnProduct_Code; } }

        public ProductOnSheet(string PrCode, int colmnPrCode)
        {
            product_Code = PrCode;
            columnProduct_Code = colmnPrCode;
        }
    }
}
