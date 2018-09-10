using OSIsoft.AF.Asset;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ilimspecs.Data
{
    interface ISpecs
    {
        List<EntityProductionUnit> GetProductUnitAll();//получение списка оборудования из БД Specs
        List<EntityProduct> GetAllProductByUnit(string unit);//Получение всех продуктов по еденице оборудования
        List<ListViewItem> GetAllProductByUnitListView(string unit);//Получение всех продуктов по имени оборудования (ProductionUnit_Desc из табл ProductionUnits БД Specs)
        List<EntitySpecification> GetAllSpecification(List<EntityVarSpecsAF> lstspecs);//Получения продуктов с границами для списка спецификаций List<AFElement>

        int SaveNewProductionUnitAndGet(string ProductionUnit_Desc);//Метод создания нового оборудования(ProductionUnit) в БД Specs
        int UpDateProductionUnit(int ProductionUnit_Id, string ProductionUnit_Desc);//Метод обновления единицы оборудования ProductionUnit

        int SaveNewProduct(string Product_Code, string Product_Desc, string ProductionUnit_Desc, string product_SAP, string product_Procont);//Создание нового продукта в табл Products в БД Specs
        int UpDateProduct(int Product_Id, string Product_Code, string Product_Desc, string ProductionUnit_Desc, string product_SAP, string product_Procont);//Обновление продуктов в таблице Products БД Specs
        VarSpecsChangeStatus UpDateVarSpecs(ProductLimit prlimit);//Обновление или создание(если их нет) границ табл VarSpecs в БД Specs по странной логике
        Int64 GetProductID(string Product_Code, string ProductionUnit_Desc);//Получение ProductID по Product_Code и ProductionUnit_Desc

        int GetCountOfProductByProductionUnitDesc(string ProductionUnit_Desc);//получение количество выпускаемых продуктов на определенной единице оборудования ProductionUnit_Desc
        List<EntityVarSpecCommonBound> GetCommonBound(string ProductionUnit_Desc);//получение всех общих границ производимых на определенной единице оборудования ProductionUnit_Desc
        List<EntityVarSpecCommonBound> GetCommonBoundCheckLimits(string ProductionUnit_Desc, CheckLimits checkLimits);//получение выбранных общих границ производимых на определенной единице оборудования ProductionUnit_Desc

        VarSpecsChangeStatus UpDateVarSpecsCommonBound(ProductLimit prlimit);//Обновление границ для всех продуктов для выбранного тега на определенном оборудовании
        List<ProductLimitCommondBound> GetProductLimitCommondBound(string Tag, string ProductionUnit_Desc);//Получение списка границ по продукту по имени тега и имени производящего оборудования

        int UpdateVarSpecsOneBound(ProductLimit prlimit, Int64 VS_ID, Int64 Product_Id);//Обновление границы по одному тегу и одному продукту
        int CreateVarSpecsOneBound(ProductLimit prlimit,  EntityProduct entityProduct);//Создание границы по одному тегу и одному продукты

    }
}
