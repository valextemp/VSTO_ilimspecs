using OSIsoft.AF;
using OSIsoft.AF.Asset;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ilimspecs.Data
{
    interface IAFEnvironment
    {
        AFDatabase GetDatabase(string server, string database);////Возвращает указатель на AF на указанном сервере
        AFDatabase GetDatabase();//Возвращает указатель на AF
        List<string> GetProductUnitFromAFUniq();//Получение уникального списка имен оборудования из ФА
        List<AFElement> GetProductUnitFromAF();//Получение списка AF элементов 
        List<ListViewItem> GetSpecificationFromAFListView(AFElement rootElement);//Получение границ (VarSpecs) выбранного элемента AF
        List<EntityVarSpecsAF> GetSpecificationAF(AFElement rootElement);//Получение границ (VarSpecs) выбранного элемента AF


    }
}
