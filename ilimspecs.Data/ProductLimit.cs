using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //Вспомогательный класс для записи в табл VarSpecs БД Specs
    public class ProductLimit
    {

        Dictionary<string, bool> Rules = new Dictionary<string, bool>();
        decimal? lolo;
        decimal? lo;
        decimal? hihi;
        decimal? hi;
        decimal? target;
        string author;

        string productionUnit_Desc; //Оборудование по которому идет запись границ (напр "ОМС")
        string tag; //Тег спецификации
        string productCode; //Код продукта для которого изменяется спецификация

        public string ProductionUnit_Desc { get => productionUnit_Desc; }
        public string Tag { get => tag; }
        public string ProductCode { get => productCode; set => productCode=value; }

        public ProductLimit(string PrUnit_Desc, string Tg, string PrCode, string auth)
        {
            Rules["LoLo"] = false;
            Rules["Lo"] = false;
            Rules["HiHi"] = false;
            Rules["Hi"] = false;
            Rules["Target"] = false;
            lolo = decimal.MaxValue;
            lo = decimal.MaxValue;
            hihi = decimal.MaxValue;
            hi = decimal.MaxValue;
            target = decimal.MaxValue;
            author = auth;
            productionUnit_Desc = PrUnit_Desc;
            tag = Tg;
            productCode = PrCode;
        }

        public string Author { get => author; set => author = value; }

        public bool LoLoRule
        {
            get { return Rules["LoLo"]; }
            set { Rules["LoLo"] = value; }
        }
        public bool LoRule
        {
            get { return Rules["Lo"]; }
            set { Rules["Lo"] = value; }
        }
        public bool HiHiRule
        {
            get { return Rules["HiHi"]; }
            set { Rules["HiHi"] = value; }
        }
        public bool HiRule
        {
            get { return Rules["Hi"]; }
            set { Rules["Hi"] = value; }
        }
        public bool TargetRule
        {
            get { return Rules["Target"]; }
            set { Rules["Target"] = value; }
        }

        public decimal? LoLo
        {
            get { return lolo; }
            set
            {
                LoLoRule = true;
                lolo = value;
            }
        }
        public decimal? Lo
        {
            get { return lo; }
            set
            {
                LoRule = true;
                lo = value;
            }
        }
        public decimal? HiHi
        {
            get { return hihi; }
            set
            {
                HiHiRule = true;
                hihi = value;
            }
        }
        public decimal? Hi
        {
            get { return hi; }
            set
            {
                HiRule = true;
                hi = value;
            }
        }
        public decimal? Target
        {
            get { return target; }
            set
            {
                TargetRule = true;
                target = value;
            }
        }

        public bool CheckLimits()
        {
            bool flagResult = false;

            decimal[] myarray = new decimal[5];
            int mycount = 0;



            if (lolo.HasValue) myarray[mycount++] = lolo.Value;
            if (lo.HasValue) myarray[mycount++] = lo.Value;
            if (target.HasValue) myarray[mycount++] = target.Value;
            if (hi.HasValue) myarray[mycount++] = hi.Value;
            if (hihi.HasValue) myarray[mycount++] = hihi.Value;

            if (mycount == 0 || mycount == 1) return true;

            for (int i = 0; i < mycount - 1; i++)
            {
                if (myarray[i] < myarray[i + 1])
                {
                    flagResult = true;
                }
                else
                {
                    flagResult = false;
                    return flagResult;
                }
            }

            return flagResult;
        }

    }
}
