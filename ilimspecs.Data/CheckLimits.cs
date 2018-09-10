using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ilimspecs.Data
{
    //Класс для записи какие границы выводить в таблицу
    public class CheckLimits
    {
        public bool ShowLoLo { get; set; }
        public bool ShowLo { get; set; }
        public bool ShowTarget { get; set; }
        public bool ShowHi { get; set; }
        public bool ShowHiHi { get; set; }
        public bool ShowAuthor { get; set; }
    }
}
