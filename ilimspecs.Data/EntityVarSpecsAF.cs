using OSIsoft.AF.Asset;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //Класс для описания спецификации в AF
    public class EntityVarSpecsAF
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Path { get; set; }
        public AFElement AFElem { get; set; }

        public override string ToString()
        {
            return Description;
        }
    }
}
