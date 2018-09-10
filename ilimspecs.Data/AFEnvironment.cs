using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using OSIsoft.AF;
using OSIsoft.AF.Search;
using OSIsoft.AF.Asset;
using System.Windows.Forms;


namespace ilimspecs.Data
{
    public class AFEnvironment: IAFEnvironment
    {
        ///<summary>
        ///_AFServer Указывает на экземпляр AF сервера
        ///</summary>  
        static string _AFServer; //Сервер AF с которым работаем в данном случае "VMILIM"
        static string _AFDB; //Имя базы данных в AF с которой работаем в данном случае "Koryazhma"
        static string _AFTree; //Ветка с которой буде работать в AF в данном случае "/Контроль"
        static string _AFCategoryProductinoUnit; //Имя категории в данном случае "Оборудование производящее продукцию" (ProductionUnit)
        static string _AFCategoryControl; //Категория "Контроль" все эти элементы являются "Спецификациями-Specification"
        static string _CodeProductionSector;//Название атрибута содержащего имя оборудования производящего оборудования
        static string _PITag;//Название атрибута в котором содержится спецификайия для VarSpecs

        public static string AFServer { get => _AFServer;  }
        public static string AFDB { get => _AFDB;  }
        public static string AFTree { get => _AFTree;  }
        public static string AFCategoryProductinoUnit { get => _AFCategoryProductinoUnit;  }
        public static string AFCategoryControl { get => _AFCategoryControl;  }
        public static string CodeProductionSector { get => _CodeProductionSector;  }
        public static string PITag { get => _PITag;  }


        //string dp = ConfigurationManager.AppSettings["provider"];
        static AFEnvironment()
        {
             _AFServer = ConfigurationManager.AppSettings["AFServer"];
             _AFDB = ConfigurationManager.AppSettings["AFDB"];
             _AFTree = ConfigurationManager.AppSettings["AFTree"];
             _AFCategoryProductinoUnit = ConfigurationManager.AppSettings["AFCategoryProductinoUnit"];
             _AFCategoryControl = ConfigurationManager.AppSettings["AFCategoryControl"];
             _CodeProductionSector = ConfigurationManager.AppSettings["ProductionUnit_Desc"];
             _PITag = ConfigurationManager.AppSettings["AttributeNameForPITag"];
        }



        public AFEnvironment()
        {
            
            //ConfigurationManager.AppSettings["MyKey"];
            //TODO: Зделать загрузуку параметров для AF из config файла
            //_AFServer = "VMILIM";
            //_AFDB = "Koryazhma";
            //_AFTree = "Контроль";
            //_AFCategoryProductinoUnit = "Оборудование производящее продукцию";
            //_AFCategoryControl = "Контроль";
            //_CodeProductionSector = "Код производственного участка";
            //_PITag = "Тэг PI";

            //_AFServer= Settings1.Default.Setting
            //_AFServer = Properties.Settings.
        }

        ///<summary>
        ///Возвращает указатель на БД AF на указанном сервере
        ///</summary>  
        public AFDatabase GetDatabase(string server, string database)
        {
            PISystems piSystems = new PISystems();
            PISystem assetServer = piSystems[server];
            AFDatabase afDatabase = assetServer.Databases[database];
            return afDatabase;
        }

        public AFDatabase GetDatabase()
        {
            PISystems piSystems = new PISystems();
            PISystem assetServer = piSystems[_AFServer];
            AFDatabase afDatabase = assetServer.Databases[_AFDB];
            return afDatabase;
        }

        /// <summary>
        /// Возвращяет уникальные "Коды производственных участков" повторения удаляются
        /// ot элементов из категории "Оборудование производящее продукцию" в ветке "Контроль"
        /// CPS -- Code Production Sector
        /// </summary>
        /// <returns></returns>
        public List<string> GetProductUnitFromAFUniq()
        {
            List<string> elements = new List<string>();

            AFDatabase database;
            string queryString;

            database = GetDatabase(_AFServer, _AFDB);
            queryString = String.Format("Root:\"{0}\" Category:\"{1}\"", _AFTree, _AFCategoryProductinoUnit);

            using (AFElementSearch elementQuery = new AFElementSearch(database, "*", queryString))
            {
                elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);

                foreach (AFElement element in elementQuery.FindElements())
                {

                    //elements.Add(element.Attributes[_CodeProductionSector].GetValue().ToString());
                    elements.Add(element.Attributes[_CodeProductionSector].GetValue().ToString().ToUpper());
                }
            }
            //var ff = elements.Distinct();
            //return ff.ToList<string>();
            return elements.Distinct().OrderBy(n => n).ToList<string>();
        }

        /// <summary>
        /// Получение списка элементов из категории "Оборудование производящее продукцию" в ветке "Контроль"
        /// </summary>
        /// <returns></returns>
        public List<AFElement> GetProductUnitFromAF()
        {
            // Здесь мы получили список элементов AF из ветки "Контроль" у которых есть категория "Оборудование производящее продукцию"
            List<AFElement> elements = new List<AFElement>();

            AFDatabase database;
            string queryString;

            database = GetDatabase(_AFServer, _AFDB);
            queryString = String.Format("Root:\"{0}\" Category:\"{1}\"", _AFTree, _AFCategoryProductinoUnit);

            using (AFElementSearch elementQuery = new AFElementSearch(database, "*", queryString))
            {
                elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);
                //elements.AddRange(elementQuery.FindElements());

                foreach (AFElement element in elementQuery.FindElements())
                {
                    elements.Add(element);
                }
            }
            return elements;
        }

        /// <summary>
        /// Метод возвращает список спецификаций относительно переданного корневого элемента
        /// </summary>
        /// <param name="rootElement">AFElement относительно которого ищутся спецификации (от него и ниже по дереву)</param>
        /// <returns>Возвращает коллекцию AFElement-ов, которые являются спецификациями</returns>
        public List<ListViewItem> GetSpecificationFromAFListView(AFElement rootElement)
        {
            if (rootElement == null) return null;

            List<ListViewItem> lestspecs = new List<ListViewItem>();

            if (rootElement.Categories[_AFCategoryControl] != null)
            {
                lestspecs.Add(new ListViewItem { Text = rootElement.Description , Tag = rootElement, ToolTipText = rootElement.Name, Checked = true });
                return lestspecs;
            }

            AFDatabase database;
            string queryString;

            database = GetDatabase(_AFServer, _AFDB);

            string myPath = rootElement.GetPath();
            queryString = String.Format("Root:\"{0}\" Category:\"{1}\"", myPath, _AFCategoryControl);

            using (AFElementSearch elementQuery = new AFElementSearch(database, "*", queryString))
            {
                elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);
                //lestspecs.AddRange(elementQuery.FindElements());

                foreach (AFElement element in elementQuery.FindElements())
                {
                    lestspecs.Add(new ListViewItem { Text = element.Description, Tag = element, ToolTipText = element.Name, Checked = true });
                }
            }
            return lestspecs;
        }

        /// <summary>
        /// Метод возвращает список спецификаций относительно переданного корневого элемента
        /// </summary>
        /// <param name="rootElement">AFElement относительно которого ищутся спецификации (от него и ниже по дереву)</param>
        /// <returns>Возвращает коллекцию EntityVarSpecsAF-ов, которые являются спецификациями</returns>
        public List<EntityVarSpecsAF> GetSpecificationAF(AFElement rootElement)
        {
            if (rootElement == null) return null;

            List<EntityVarSpecsAF> lestspecs = new List<EntityVarSpecsAF>();

            if (rootElement.Categories[_AFCategoryControl] != null)
            {
                string myPath1 = rootElement.Parent.GetPath();
                lestspecs.Add(new EntityVarSpecsAF { Description = rootElement.Description, Path = myPath1.Substring(myPath1.IndexOf(AFEnvironment.AFTree)), Name = rootElement.Name, AFElem=rootElement});
                return lestspecs;
            }

            AFDatabase database;
            string queryString;

            database = GetDatabase(_AFServer, _AFDB);

            string myPath = rootElement.GetPath();
            queryString = String.Format("Root:\"{0}\" Category:\"{1}\"", myPath, _AFCategoryControl);

            using (AFElementSearch elementQuery = new AFElementSearch(database, "*", queryString))
            {
                elementQuery.CacheTimeout = TimeSpan.FromMinutes(5);
                //lestspecs.AddRange(elementQuery.FindElements());
                string myPath2;
                foreach (AFElement element in elementQuery.FindElements())
                {
                    myPath2 = element.Parent.GetPath();
                    lestspecs.Add(new EntityVarSpecsAF { Description = element.Description, Path = myPath2.Substring(myPath2.IndexOf(AFEnvironment.AFTree)), Name = element.Name, AFElem=element });
                }
            }
            return lestspecs;
        }

        


    }
}
