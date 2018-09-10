using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ilimspecs.Data;

namespace ilimspecs.BL
{
    interface IExcelBook
    {
        void CleanSheet(); //Очитска страницы перед выводом

        void ProductionUnitTitle(); //Вывод заголовка для оборудования (табл ProductionUnit)
        void ProductionUnitTablePrint();//Вывод таблички ProductionUnit-ов вместе с заголовком
        bool ProductionUnitCheckSheet();//Проверка перед записью в Specs, что активный лист является листом с оборудованием

        void ProductTitle();//Вывод заголовка таблицы с продуктами
        void ProductTablePrint(List<EntityProduct> lstproduct);//Вывод таблички Producti-ов вместе с заголовком
        bool ProductCheckSheet();//Проверка перед записью в Specs, что активный лист является листом с продуктами

        void VarSpecsTitle(List<string> lstproduct, CheckLimits checkLimits);//Вывод заголовка таблицы со спецификациями (VarSpecs)
        void VarSpecsTablePrint(List<EntityVarSpecsAF> listvarspecs, List<string> lstproduct, CheckLimits checkLimits);//Вывод таблицы со спецификациями (VarSpecs)
        bool VarSpecsCheckSheet(); //Проверка перед записью в Specs, что активный лист является листом со спецификациями

        void VarSpecsCommonLimitsTitle(CheckLimits checkLimits);//Вывод заголовка с общими границами по всем продуктам
        void VarSpecsCommonLimitsTablePrint(List<EntityVarSpecsAF> listvarspecs, CheckLimits checkLimits);//Вывод содержимого со спецификациями (внутри вызывается метод вывода заголовка таблицы)
        bool VarSpecsCommonLimitsCheckSheet();//Проверка перед записью в Specs, что активный лист является листом с общими границами

        void ProductionUnitSaveFromSheet();//Метод создания или обновления оборудования в БД Specs
        void ProductSaveFromSheet();//Метод создания или обновления продуктов в БД Specs
        void VarSpecsSaveFromSheet();//Метод создания или обновления границ по продуктам для спецификаций
        void VarSpecsCommonLimitSaveFromSheet();//Метод создания или обновления границ по всем продуктам

        List<EntityProductionUnit> ReturnCPS_AF_Specs(); //Возвращение списка EntityProductionUnit (из AF с учетом тех которые есть в specs)
    }
}
