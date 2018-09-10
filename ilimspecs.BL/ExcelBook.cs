using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ilimspecs.Data;
using System.Windows.Forms;
using System.Drawing;

namespace ilimspecs.BL
{
    public enum MarkType { Font = 0, Interior = 1 }

    public class ExcelBook: IExcelBook
    {
        Specs specs ;

        Workbook workbook;
        Worksheet worksheet;
        Microsoft.Office.Interop.Excel.Application excelApplication;

        //Поле показывающее как выделять ячейки (или менять цвет шрифта=0 или менять заливку ячейки=1)
        MarkType mark = MarkType.Font;
        public MarkType Mark
        {
            set
            {
                switch (value)
                {
                    case MarkType.Font:
                        mark = MarkType.Font;
                        break;
                    case MarkType.Interior:
                        mark = MarkType.Interior;
                        break;
                    default:
                        mark = MarkType.Font;
                        break;
                }
            }
            get
            {
                return mark;
            }
        }

        //Список для учета заголовков продуктов на листе
        List<ProductOnSheet> ListProductOnSheet=new List<ProductOnSheet>();

        public ExcelBook(Microsoft.Office.Interop.Excel.Application exlApplication, Workbook wrkbook, Worksheet wrksheet )
        {
            excelApplication = exlApplication;
            workbook = wrkbook;
            worksheet = wrksheet;
            mark = MarkType.Interior;
            specs = new Specs();

            string markType= ConfigurationManager.AppSettings["MarkType"];

            if (markType!=null)
            {
                try
                {
                    int mm = Convert.ToInt32(markType);
                    if (mm==1)
                    {
                        mark = MarkType.Interior;
                    }
                    else
                    {
                        mark = MarkType.Font;
                    }
                }
                catch (Exception)
                {
                    mark = MarkType.Font;
                }
            }
            

            
        }

        public void CleanSheet()
        {
            worksheet.Cells.Clear();
        }

        public void ProductionUnitTablePrint()
        {
            CleanSheet();//Очищаем страницу
            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            excelApplication.ScreenUpdating = false;
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            List<EntityProductionUnit> pulist = new List<EntityProductionUnit>();
            pulist = ReturnCPS_AF_Specs();
            if (pulist == null)
            {
                excelApplication.ScreenUpdating = true;
                worksheet.Application.Cursor = initialCursor;
                return; //если список оборудования пустой просто выходим
            }
            
            ProductionUnitTitle();//Печатаем заголовок таблицы

            int row = 2;
            foreach (var item in pulist)
            {
                worksheet.Cells[row, 1].Value = "x";
                //Выводим ProductionUnit_Id если оно не равно 0
                if (item.ProductionUnit_Id != 0) worksheet.Cells[row, 2].Value = item.ProductionUnit_Id;
                worksheet.Cells[row, 3].Value = item.ProductionUnit_Desc;
                row++;
            }

            worksheet.UsedRange.Columns.AutoFit();

            excelApplication.ScreenUpdating = true;
            worksheet.Application.Cursor = initialCursor;
        }

        public void ProductionUnitTitle()
        {
            worksheet.Cells[1, 1] = "Отметка";
            worksheet.Cells[1, 2] = "ID оборудования";
            worksheet.Cells[1, 3] = "Оборудование";
            worksheet.Range["A1:C1"].Font.Bold = true;
            worksheet.Range["A1:C1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Range["A1:C1"].VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        public void ProductTitle()
        {
            worksheet.Cells[1, 1] = "Отметка";
            worksheet.Cells[1, 2] = "ID продукта";
            worksheet.Cells[1, 3] = "Код продукта";
            worksheet.Cells[1, 4] = "Наименование продукта";
            worksheet.Cells[1, 5] = "Оборудование";
            worksheet.Cells[1, 6] = "Продукт SAP";
            worksheet.Cells[1, 7] = "Продукт Procont";

            worksheet.Range["A1:G1"].Font.Bold = true;
            worksheet.Range["A1:G1"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Range["A1:G1"].VerticalAlignment = XlVAlign.xlVAlignCenter;
        }


        public void ProductTablePrint(List<EntityProduct> lstproduct)
        {
            if (lstproduct == null) return;

            CleanSheet();//Очищаем страницу
            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            excelApplication.ScreenUpdating = false;
            worksheet.Application.Cursor = XlMousePointer.xlWait;
            ProductTitle();

            int row = 2;
            foreach (EntityProduct item in lstproduct)
            {
                worksheet.Cells[row, 1].Value = "x";
                worksheet.Cells[row, 2].Value = item.Product_Id;
                worksheet.Cells[row, 3].Value = item.Product_Code;
                worksheet.Cells[row, 4].Value = item.Product_Desc;
                worksheet.Cells[row, 5].Value = item.ProductionUnit_Desc;
                worksheet.Cells[row, 6].Value = item.Product_SAP;
                worksheet.Cells[row, 7].Value = item.Product_Procont;

                row++;
            }
            worksheet.UsedRange.Columns.AutoFit();

            excelApplication.ScreenUpdating = true;
            worksheet.Application.Cursor = initialCursor;
        }



        public List<EntityProductionUnit> ReturnCPS_AF_Specs()
        {
            AFEnvironment af = new AFEnvironment();
            //Specs specs = new Specs();

            List<EntityProductionUnit> resultcompare = new List<EntityProductionUnit>();

            List<string> CPSUniq = af.GetProductUnitFromAFUniq();
            List<EntityProductionUnit> CPSSpecs = specs.GetProductUnitAll();

            var resultcompare1 =
                 from cpsuniq in CPSUniq
                 join cspspecs in CPSSpecs on cpsuniq equals cspspecs.ProductionUnit_Desc into tempGroup
                 from item in tempGroup.DefaultIfEmpty(new EntityProductionUnit { ProductionUnit_Id = 0, ProductionUnit_Desc = cpsuniq })
                 select new EntityProductionUnit { ProductionUnit_Id = item.ProductionUnit_Id, ProductionUnit_Desc = cpsuniq };

            resultcompare = resultcompare1.ToList<EntityProductionUnit>();

            return resultcompare;

        }

        public void VarSpecsTitle(List<string> lstproduct, CheckLimits checkLimits)
        {

            worksheet.Cells[1, 1] = "Отметка";
            //Microsoft.Office.Interop.Excel.Range xmark = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]].Merge();
            //worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]].Merge();
            Range cellmark1 = worksheet.Cells[1, 1];
            Range cellmark2 = worksheet.Cells[2, 1];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;
            

            //Range mycell = worksheet.Cells[1, 6 + col];


            worksheet.Cells[1, 2] = "Оборудование";
            //Microsoft.Office.Interop.Excel.Range unit = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]].Merge();
            //worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]].Merge();
            cellmark1 = worksheet.Cells[1, 2];
            cellmark2 = worksheet.Cells[2, 2];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            //unit.Value = "Оборудование";

            //Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
            worksheet.Cells[1, 3] = "Путь";
            //worksheet.Range[worksheet.Cells[1, 3], worksheet.Cells[2, 3]].Merge();
            cellmark1 = worksheet.Cells[1, 3];
            cellmark2 = worksheet.Cells[2, 3];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            worksheet.Cells[1, 4] = "Имя";
            //worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[2, 4]].Merge();
            cellmark1 = worksheet.Cells[1, 4];
            cellmark2 = worksheet.Cells[2, 4];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            worksheet.Cells[1, 5] = "Тег";
            //worksheet.Range[worksheet.Cells[1, 5], worksheet.Cells[2, 5]].Merge();
            cellmark1 = worksheet.Cells[1, 5];
            cellmark2 = worksheet.Cells[2, 5];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            //worksheet.Range["A1:R1"].Font.Bold = true;
            //worksheet.Range["A1:R1"].HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            //worksheet.Range["A1:R1"].VerticalAlignment = Microsoft.Office.Core.XlVAlign.xlVAlignCenter;

            int col = 0;
            //Очищаем список продуктов на листе
            ListProductOnSheet.Clear();
            int colProduct = 0;
            foreach (var item in lstproduct)
            {
                        Range mycell = worksheet.Cells[1, 6 + col];
                        colProduct = 6 + col;
                        worksheet.Cells[1, colProduct] = item;
                        worksheet.Cells[1, colProduct].Font.Bold = true;
                        worksheet.Cells[1, colProduct].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        worksheet.Cells[1, colProduct].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        

                //Добавляю в список продуктов на листе
                ListProductOnSheet.Add(new ProductOnSheet(item, colProduct));
                        if (checkLimits.ShowLoLo)
                        {
                            worksheet.Cells[2, 6 + col] = "LoLo";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment =  Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        if (checkLimits.ShowLo)
                        {
                            worksheet.Cells[2, 6 + col] = "Lo";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        if (checkLimits.ShowTarget)
                        {
                            worksheet.Cells[2, 6 + col] = "Target";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        if (checkLimits.ShowHi)
                        {
                            worksheet.Cells[2, 6 + col] = "Hi";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        if (checkLimits.ShowHiHi)
                        {
                            worksheet.Cells[2, 6 + col] = "HiHi";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        if (checkLimits.ShowAuthor)
                        {
                            worksheet.Cells[2, 6 + col] = "Author";
                            worksheet.Cells[2, 6 + col].Font.Bold = true;
                            worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            col++;
                        }
                        Range mycell2 = worksheet.Cells[1, 6 + col - 1];
                        worksheet.Range[mycell, mycell2].Merge();
                        worksheet.Range[mycell, mycell2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        worksheet.Range[mycell, mycell2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            }
        }
        
        /// <summary>
        /// Печатает на активном листе спецификации с продуктами
        /// </summary>
        /// <param name="varspecswithproduct">Список спецификаций которые выбраны в checklistbox</param>
        /// <param name="lstproduct">Список с кодами продуктов Product_Code</param>
        /// <param name="checkLimits">Класс который содержит в себе какие границы выводить</param>
        public void VarSpecsTablePrint(List<EntityVarSpecsAF> listvarspecs, List<string> lstproduct, CheckLimits checkLimits)
        {
            CleanSheet();//Очищаем страницу

            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            excelApplication.ScreenUpdating = false;
            excelApplication.DisplayAlerts = false;
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            List<EntitySpecification> listEntitySpecification = specs.GetAllSpecification(listvarspecs);

            if (listEntitySpecification==null || listEntitySpecification.Count<=0)
            {
                excelApplication.DisplayAlerts = true;
                excelApplication.ScreenUpdating = true;
                worksheet.Application.Cursor = initialCursor;
                MessageBox.Show("Отсутствуют данные в Specs", "Сообщение об ошибке", MessageBoxButtons.OK);
                return;
            }

            VarSpecsTitle(lstproduct, checkLimits);

            
            #region Вывод таблицы с данными на лист
            // Вывод содержимого в таблицу спецификаций
            int row = 3;
            foreach (var item in listEntitySpecification)
            {
                worksheet.Cells[row, 1].Value = "x";
                worksheet.Cells[row, 2].Value = item.ProductionUnit;
                worksheet.Cells[row, 3].Value = item.PathAF;
                worksheet.Cells[row, 4] = item.Name;
                worksheet.Cells[row, 5] = item.Tag;

                foreach (ProductLimits pl in item.ProductLimits)
                {
                    //                    ProductOnSheet possheet = ListProductOnSheet.Find(x => x.Product_Code.Contains(pl.Product_Code));
                   
                    ProductOnSheet possheet = ListProductOnSheet.Find(x => String.Compare(x.Product_Code,pl.Product_Code,true)==0);
                    if (possheet != null)
                    {
                        int col = 0;
                        if (checkLimits.ShowLoLo)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.LoLo;
                            col++;
                        }
                        if (checkLimits.ShowLo)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.Lo;
                            col++;
                        }
                        if (checkLimits.ShowTarget)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.Target;
                            col++;
                        }
                        if (checkLimits.ShowHi)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.Hi;
                            col++;
                        }
                        if (checkLimits.ShowHiHi)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.HiHi;
                            col++;
                        }
                        if (checkLimits.ShowAuthor)
                        {
                            worksheet.Cells[row, possheet.ColumnProduct_Code + col] = pl.Author;
                            col++;
                        }
                    }
                }
                row++;
            }
            #endregion

            //worksheet.UsedRange.Columns.AutoFit();
            worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[worksheet.UsedRange.Rows.Count, 5]].Columns.AutoFit();

            excelApplication.DisplayAlerts = true;
            excelApplication.ScreenUpdating = true;
            worksheet.Application.Cursor = initialCursor;
        }

        public bool ProductionUnitCheckSheet()
        {
            return (worksheet.Cells[1, 1].Value == "Отметка" && worksheet.Cells[1, 2].Value == "ID оборудования" && worksheet.Cells[1, 3].Value == "Оборудование");
        }

        public bool ProductCheckSheet()
        {
            return (worksheet.Cells[1, 1].Value == "Отметка" && worksheet.Cells[1, 2].Value == "ID продукта" && worksheet.Cells[1, 3].Value == "Код продукта" && worksheet.Cells[1, 4].Value == "Наименование продукта" && worksheet.Cells[1, 5].Value == "Оборудование");
        }

        public bool VarSpecsCheckSheet()
        {
            return (worksheet.Cells[1, 1].Value == "Отметка" && worksheet.Cells[1, 2].Value == "Оборудование" && worksheet.Cells[1, 3].Value == "Путь" && worksheet.Cells[1, 4].Value == "Имя" && worksheet.Cells[1, 5].Value == "Тег"&& worksheet.Cells[1, 6].Value != "Все продукты");
        }

        public void ProductionUnitSaveFromSheet()
        {
            //Specs specs = new Specs();

            int productionUnit_Id;
            string productionUnit_Desc;
            int resultquery = 0;

            dynamic dynproductionUnit_Id;
            dynamic dynproductionUnit_Desc;

            int myrows = worksheet.UsedRange.Rows.Count;
            int mycolumns = worksheet.UsedRange.Columns.Count;

           
            for (int i = 1; i < myrows; i++)
            {

                if (worksheet.Cells[i + 1, 1].Value == "x")
                {
                    //Если все true тогда запускаем процедуру обновления или создания единицы оборудования
                    if (ProductionUnitCheckSheet())
                    {
                        dynproductionUnit_Id = worksheet.Cells[i + 1, 2].Value;
                        dynproductionUnit_Desc = worksheet.Cells[i + 1, 3].Value;
                        if ((dynproductionUnit_Id is double) && (dynproductionUnit_Desc is string))
                        {
                            //Запуск ОБНОВЛЕНИЯ единицы оборудования
                            productionUnit_Id = (int)dynproductionUnit_Id;
                            productionUnit_Desc = (string)dynproductionUnit_Desc;
                            resultquery = specs.UpDateProductionUnit(productionUnit_Id, productionUnit_Desc.ToUpper().Trim());
                            if (resultquery!=0)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], true);
                            }
                            else
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], false);
                            }
                        }
                        else if ((dynproductionUnit_Id == null) && (dynproductionUnit_Desc is string))
                        {
                            //Запуск создания единицы оборудования
                            productionUnit_Desc = (string)dynproductionUnit_Desc;
                            resultquery = specs.SaveNewProductionUnitAndGet(productionUnit_Desc.Trim());
                            if (resultquery != 0)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], true);
                                worksheet.Cells[i + 1, 2] = resultquery;
                            }
                            else
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], false);
                            }
                        }
                        else
                        {
                            MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], false);
                        }
                    }
                    else
                        {
                        MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 3]], false);
                        }
                    }
                }
            }

        public void ProductSaveFromSheet()
        {
            int product_Id;
            string product_Code;
            string product_Desc;
            string productionUnit_Desc;
            string product_SAP;
            string product_Procont;

            int resultquery;

            dynamic dynproduct_Id;
            dynamic dynproduct_Code;
            dynamic dynproduct_Desc;
            dynamic dynproductionUnit_Desc;
            dynamic dynproduct_SAP;
            dynamic dynproductProcont;




            int myrows = worksheet.UsedRange.Rows.Count;
            int mycolumns = worksheet.UsedRange.Columns.Count;

            if (ProductCheckSheet())
            {
                //Specs specs = new Specs();
                for (int i = 1; i < myrows; i++)
                {
                    if (worksheet.Cells[i + 1, 1].Value == "x")
                    {
                        //Если все true тогда запускаем процедуру создания или обновления
                        dynproduct_Id = worksheet.Cells[i + 1, 2].Value;
                        dynproduct_Code = worksheet.Cells[i + 1, 3].Value;
                        dynproduct_Desc = worksheet.Cells[i + 1, 4].Value;
                        dynproductionUnit_Desc = worksheet.Cells[i + 1, 5].Value;
                        dynproduct_SAP = worksheet.Cells[i + 1, 6].Value;
                        dynproductProcont = worksheet.Cells[i + 1, 7].Value;


                        if ((dynproduct_Id == null) && (dynproduct_Code is string))
                        {
                            //Запуск создания продукта
                            product_Code = (string)dynproduct_Code;
                            product_Desc = (string)dynproduct_Desc;
                            productionUnit_Desc = (string)dynproductionUnit_Desc;
                            if ((string)dynproduct_SAP != null) product_SAP = (string)dynproduct_SAP;
                            else product_SAP = "";
                            if ((string)dynproductProcont != null) product_Procont = (string)dynproductProcont;
                            else product_Procont = "";

                            resultquery = specs.SaveNewProduct(product_Code, product_Desc, productionUnit_Desc, product_SAP, product_Procont);
                            if (resultquery != 0)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]], true);
                                worksheet.Cells[i + 1, 2].Value = resultquery;
                            }
                            else
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]], false);
                            }
                        }
                        else if ((dynproduct_Id is double) && (dynproduct_Code is string))
                        {
                            product_Id = (int)dynproduct_Id;
                            if ((string)dynproduct_Code != null) product_Code = (string)dynproduct_Code;
                            else product_Code = "";
                            if ((string)dynproduct_Desc != null) product_Desc = (string)dynproduct_Desc;
                            else product_Desc = "";
                            if ((string)dynproductionUnit_Desc != null) productionUnit_Desc = (string)dynproductionUnit_Desc;
                            else productionUnit_Desc = "";
                            if ((string)dynproduct_SAP != null) product_SAP = (string)dynproduct_SAP;
                            else product_SAP = "";
                            if ((string)dynproductProcont != null) product_Procont = (string)dynproductProcont;
                            else product_Procont = "";

                            //Запуск обновления продукта
                            resultquery = specs.UpDateProduct(product_Id, product_Code.Trim(), product_Desc.Trim(), productionUnit_Desc.Trim(), product_SAP.Trim(), product_Procont.Trim());

                            if (resultquery != 0)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]], true);
                            }
                            else
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]], false);
                            }

                        }

                        else
                        {
                            MarkRange(worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]], false);
                        }
                    }
                    else
                    {
                        worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        //worksheet.Range[worksheet.Cells[i + 1, 1], worksheet.Cells[i + 1, 7]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    }
                }
            }
            //
            else
            {
                MessageBox.Show("Данные на листе не соответствуют выбранной операции", "Сообщение об ошибке", MessageBoxButtons.OK);
            }

        }
       

        /// <summary>
        /// Метод сохранения границ по продуктам с листа
        /// </summary>
        public void VarSpecsSaveFromSheet()
        {
            //Specs specs = new Specs();

            dynamic dynProductionUnit_Desc;
            dynamic dynTag;
            dynamic dynProductCode;
            dynamic dynPathAF;
            dynamic dynName;


            dynamic dynLoLo;
            dynamic dynLo;
            dynamic dynTarget;
            dynamic dynHiHi;
            dynamic dynHi;
            dynamic dynAuthor;

            VarSpecsChangeStatus resultquery;

            int myrows = worksheet.UsedRange.Rows.Count;
            int mycolumns = worksheet.UsedRange.Columns.Count;

            if (VarSpecsCheckSheet())
            {
                //Определяем автора изменяющего значения границ
                string author = string.Format(Environment.UserDomainName + "/" + Environment.UserName);


                // Количество границ которые нужно записать
                int SpecsLimits = worksheet.Cells[1, 6].MergeArea.Columns.Count;

                string ProductionUnit_Desc;
                string Tag;
                string ProductCode;
                List<ProductLimit> listlimit = new List<ProductLimit>();


                dynamic dynTemp;
                //    //С тройки начинается потому-что заголовок таблицы занимает 2 строки
                for (int rowcount = 3; rowcount < myrows + 1; rowcount++)
                {
                    if (worksheet.Cells[rowcount, 1].Value == "x")
                    {
                        //Получаем значения ProductionUnit_Desc, Tag
                        dynProductionUnit_Desc = worksheet.Cells[rowcount, 2].Value;
                        dynTag = worksheet.Cells[rowcount, 5].Value;

                        ProductionUnit_Desc = (string)dynProductionUnit_Desc;
                        Tag = (string)dynTag;

                        for (int columncount = 6; columncount < mycolumns; columncount += SpecsLimits)
                        {
                            // Получаю код продукта ProductCode
                            dynProductCode = worksheet.Cells[1, columncount].Value;
                            ProductCode = (string)dynProductCode;
                            ProductLimit prlimit = new ProductLimit(ProductionUnit_Desc, Tag, ProductCode, author);


                            //Здесь пробегаемся по границам продуктов
                            for (int limitsCount = 0; limitsCount < SpecsLimits; limitsCount++)
                            {

                                if ((string)worksheet.Cells[2, columncount + limitsCount].Value == "LoLo")
                                {
                                    dynLoLo = worksheet.Cells[rowcount, columncount + limitsCount].Value;

                                    if ((dynLoLo is double) || (dynLoLo == null))
                                    {
                                        if (dynLoLo == null)
                                        {
                                            prlimit.LoLo = null;
                                        }
                                        else
                                        {
                                            prlimit.LoLo = (decimal)dynLoLo;
                                        }
                                        continue;
                                    }
                                    //if ((dynLoLo is double) && (dynproduct_Code is string))
                                    //if ((dynproduct_Id == null) && (dynproduct_Code is string))

                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount+ myrows+3, columncount + limitsCount].Value=dynLoLo;
                                }
                                if ((string)worksheet.Cells[2, columncount + limitsCount].Value == "Lo")
                                {
                                    dynLo = worksheet.Cells[rowcount, columncount + limitsCount].Value;
                                    if ((dynLo is double) || (dynLo == null))
                                    {
                                        if (dynLo == null)
                                        {
                                            prlimit.Lo = null;
                                        }
                                        else
                                        {
                                            prlimit.Lo = (decimal)dynLo;
                                        }
                                        continue;
                                    }
                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value =dynLo;
                                }
                                if ((string)worksheet.Cells[2, columncount + limitsCount].Value == "Target")
                                {
                                    dynTarget = worksheet.Cells[rowcount, columncount + limitsCount].Value;
                                    if ((dynTarget is double) || (dynTarget == null))
                                    {
                                        if (dynTarget == null)
                                        {
                                            prlimit.Target = null;
                                        }
                                        else
                                        {
                                            prlimit.Target = (decimal)dynTarget;
                                        }
                                        continue;
                                    }
                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value =dynTarget;
                                }
                                if ((string)worksheet.Cells[2, columncount + limitsCount].Value == "HiHi")
                                {
                                    dynHiHi = worksheet.Cells[rowcount, columncount + limitsCount].Value;
                                    if ((dynHiHi is double) || (dynHiHi == null))
                                    {
                                        if (dynHiHi == null)
                                        {
                                            prlimit.HiHi = null;
                                        }
                                        else
                                        {
                                            prlimit.HiHi = (decimal)dynHiHi;
                                        }
                                        continue;
                                    }

                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value = dynHiHi;
                                }
                                if ((string)worksheet.Cells[2, columncount + limitsCount].Value == "Hi")
                                {
                                    dynHi = worksheet.Cells[rowcount, columncount + limitsCount].Value;

                                    if ((dynHi is double) || (dynHi == null))
                                    {
                                        if (dynHi == null)
                                        {
                                            prlimit.Hi = null;
                                        }
                                        else
                                        {
                                            prlimit.Hi = (decimal)dynHi;
                                        }
                                        continue;
                                    }
                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value = dynHi;
                                }

                            }

                            // !!!!! Здесь нужно записать Отдельный экземпляр ProductLimits В БД Specs
                            //listlimit.Add(prlimit);

                            resultquery = specs.UpDateVarSpecs(prlimit);

                        if (resultquery == VarSpecsChangeStatus.ChangeOK)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[rowcount, columncount], worksheet.Cells[rowcount, columncount + SpecsLimits - 1]], true);
                            }
                            else if (resultquery == VarSpecsChangeStatus.ChangeBad)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[rowcount, columncount], worksheet.Cells[rowcount, columncount + SpecsLimits - 1]], false);
                            }
                        }

                    }
                }

            }
            else
            {
                //Данные на листе не соответствуют выбранной операции
                MessageBox.Show("Данные на листе не соответствуют выбранной операции", "Сообщение об ошибке", MessageBoxButtons.OK);
            }

        }

        /// <summary>
        /// Метод для выделения диапазона цветом 
        /// </summary>
        /// <param name="range">Диапазон который нужно выделить</param>
        /// <param name="status">Если true выделяем зеленым, false выделяем красным</param>
        private void MarkRange(Range range, bool status)
        {
            if (status)
            {
                if (mark==MarkType.Font)
                {
                    range.Font.Color= System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                }
                else
                {
                    range.Interior.Color= System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                }
            }
            else
            {
                if (mark == MarkType.Font)
                {
                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
                else
                {
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                }
            }
        }

        public void VarSpecsCommonLimitsTitle(CheckLimits checkLimits)
        {
            worksheet.Cells[1, 1] = "Отметка";
            //Microsoft.Office.Interop.Excel.Range xmark = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]].Merge();
            //worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]].Merge();
            Range cellmark1 = worksheet.Cells[1, 1];
            Range cellmark2 = worksheet.Cells[2, 1];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            //Range mycell = worksheet.Cells[1, 6 + col];


            worksheet.Cells[1, 2] = "Оборудование";
            //Microsoft.Office.Interop.Excel.Range unit = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]].Merge();
            //worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]].Merge();
            cellmark1 = worksheet.Cells[1, 2];
            cellmark2 = worksheet.Cells[2, 2];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;
            //unit.Value = "Оборудование";

            //Excel.Range range1 = sheet.get_Range(sheet.Cells[1, 1], sheet.Cells[9, 9]);
            worksheet.Cells[1, 3] = "Путь";
            //worksheet.Range[worksheet.Cells[1, 3], worksheet.Cells[2, 3]].Merge();
            cellmark1 = worksheet.Cells[1, 3];
            cellmark2 = worksheet.Cells[2, 3];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            worksheet.Cells[1, 4] = "Имя";
            //worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[2, 4]].Merge();
            cellmark1 = worksheet.Cells[1, 4];
            cellmark2 = worksheet.Cells[2, 4];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            worksheet.Cells[1, 5] = "Тег";
            //worksheet.Range[worksheet.Cells[1, 5], worksheet.Cells[2, 5]].Merge();
            cellmark1 = worksheet.Cells[1, 5];
            cellmark2 = worksheet.Cells[2, 5];
            worksheet.Range[cellmark1, cellmark2].Merge();
            worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Range[cellmark1, cellmark2].Font.Bold = true;

            worksheet.Cells[1, 6] = "Все продукты";
            worksheet.Cells[1, 6].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[1, 6].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            worksheet.Cells[1, 6].Font.Bold = true;

            //worksheet.Range[worksheet.Cells[1, 5], worksheet.Cells[2, 5]].Merge();
            //cellmark1 = worksheet.Cells[1, 5];
            //cellmark2 = worksheet.Cells[2, 5];
            //worksheet.Range[cellmark1, cellmark2].Merge();
            //worksheet.Range[cellmark1, cellmark2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            //worksheet.Range[cellmark1, cellmark2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            //worksheet.Range["A1:R1"].Font.Bold = true;
            //worksheet.Range["A1:R1"].HorizontalAlignment = Microsoft.Office.Core.XlHAlign.xlHAlignCenter;
            //worksheet.Range["A1:R1"].VerticalAlignment = Microsoft.Office.Core.XlVAlign.xlVAlignCenter;

            int col = 0;
            int colProduct = 0;

            Range mycell = worksheet.Cells[1, 6 + col];
                colProduct = 6 + col;


            if (checkLimits.ShowLoLo)
                {
                    worksheet.Cells[2, 6 + col] = "LoLo";
                    worksheet.Cells[2, 6 + col].Font.Bold = true;
                    worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    col++;
                }
                if (checkLimits.ShowLo)
                {
                    worksheet.Cells[2, 6 + col] = "Lo";
                    worksheet.Cells[2, 6 + col].Font.Bold = true;
                    worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    col++;
            }
            if (checkLimits.ShowTarget)
                {
                    worksheet.Cells[2, 6 + col] = "Target";
                    worksheet.Cells[2, 6 + col].Font.Bold = true;
                    worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    col++;
            }
            if (checkLimits.ShowHi)
                {
                    worksheet.Cells[2, 6 + col] = "Hi";
                    worksheet.Cells[2, 6 + col].Font.Bold = true;
                    worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    col++;
            }
            if (checkLimits.ShowHiHi)
                {
                    worksheet.Cells[2, 6 + col] = "HiHi";
                    worksheet.Cells[2, 6 + col].Font.Bold = true;
                    worksheet.Cells[2, 6 + col].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Cells[2, 6 + col].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                    col++;
            }
            Range mycell2 = worksheet.Cells[1, 6 + col - 1];
                worksheet.Range[mycell, mycell2].Merge();
                worksheet.Range[mycell, mycell2].VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Range[mycell, mycell2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        /// <summary>
        /// Печатает таблицу с общими границами для спецификаций
        /// </summary>
        /// <param name="listvarspecs">Список спецификаций</param>
        /// <param name="checkLimits">Перечень выводимых полей</param>
        public void VarSpecsCommonLimitsTablePrint(List<EntityVarSpecsAF> listvarspecs, CheckLimits checkLimits)
        {
            if (listvarspecs==null || listvarspecs.Count==0)
            {
                //Ошибка 
                return;
            }
            CleanSheet();
            EntityVarSpecsAF entityVarSpecsAF = listvarspecs[0];
            string unit = entityVarSpecsAF.AFElem.Attributes[AFEnvironment.CodeProductionSector].GetValue().ToString();

            //Это вызов метода который выдает все границы с одинаковыми цифрами
            //List<EntityVarSpecCommonBound> listCommonBound = specs.GetCommonBound(unit);


            List<EntityVarSpecCommonBound> listCommonBound = specs.GetCommonBoundCheckLimits(unit, checkLimits);

            //for (int i = 0; i < listvarspecs.Count; i++)
            int row = 3;
            int myColumn = 0;

            excelApplication.ScreenUpdating = false;
            excelApplication.DisplayAlerts = false;
            VarSpecsCommonLimitsTitle(checkLimits);

            //Будем искать строки без учета регистра
            
            foreach (EntityVarSpecsAF item in listvarspecs)
            {
                worksheet.Cells[row, 1].Value = "x";
                worksheet.Cells[row, 2].Value = item.AFElem.Attributes[AFEnvironment.CodeProductionSector].GetValue().ToString();
                worksheet.Cells[row, 3].Value = item.Path;
                worksheet.Cells[row, 4] = item.Name;
                worksheet.Cells[row, 5] = item.AFElem.Attributes[AFEnvironment.PITag].GetValue().ToString();

                
                //Если список с общими границацами из БД Specs не  NULL то надо вывести значения для этих тегов
                if (listCommonBound!=null)
                {
                    //                    EntityVarSpecCommonBound entityVarSpecCommonBound = listCommonBound.Find(x => x.Var_Id.Contains(item.AFElem.Attributes[AFEnvironment.PITag].GetValue().ToString(), compare));
                    EntityVarSpecCommonBound entityVarSpecCommonBound = listCommonBound.Find(x =>String.Compare(x.Var_Id,item.AFElem.Attributes[AFEnvironment.PITag].GetValue().ToString(),true)==0);
                    if (entityVarSpecCommonBound!=null)
                    {
                        int col = 0;
                        if (checkLimits.ShowLoLo)
                        {
                            worksheet.Cells[row, 6 + col] = entityVarSpecCommonBound.LoLo;
                            col++;
                        }
                        if (checkLimits.ShowLo)
                        {
                            worksheet.Cells[row, 6 + col] = entityVarSpecCommonBound.Lo;
                            col++;
                        }
                        if (checkLimits.ShowTarget)
                        {
                            worksheet.Cells[row, 6 + col] = entityVarSpecCommonBound.Target;
                            col++;
                        }
                        if (checkLimits.ShowHi)
                        {
                            worksheet.Cells[row, 6 + col] = entityVarSpecCommonBound.Hi;
                            col++;
                        }
                        if (checkLimits.ShowHiHi)
                        {
                            worksheet.Cells[row, 6 + col] = entityVarSpecCommonBound.HiHi;
                            col++;
                        }
                            //worksheet.Cells[row, 6+ col] = "Общая граница";
                            //col++;
                        myColumn = col;
                    }
                    else
                    {
                        int col = 0;
                        if (checkLimits.ShowLoLo)
                        {
                            worksheet.Cells[row, 6 + col] = "Разные";
                            col++;
                        }
                        if (checkLimits.ShowLo)
                        {
                            worksheet.Cells[row, 6 + col] = "Разные";
                            col++;
                        }
                        if (checkLimits.ShowTarget)
                        {
                            worksheet.Cells[row, 6 + col] = "Разные";
                            col++;
                        }
                        if (checkLimits.ShowHi)
                        {
                            worksheet.Cells[row, 6 + col] = "Разные";
                            col++;
                        }
                        if (checkLimits.ShowHiHi)
                        {
                            worksheet.Cells[row, 6 + col] = "Разные";
                            col++;
                        }
                        myColumn = col;
                    }
                }
                else
                {
                    int col = 0;
                    if (checkLimits.ShowLoLo)
                    {
                        worksheet.Cells[row, 6 + col] = "Разные";
                        col++;
                    }
                    if (checkLimits.ShowLo)
                    {
                        worksheet.Cells[row, 6 + col] = "Разные";
                        col++;
                    }
                    if (checkLimits.ShowTarget)
                    {
                        worksheet.Cells[row, 6 + col] = "Разные";
                        col++;
                    }
                    if (checkLimits.ShowHi)
                    {
                        worksheet.Cells[row, 6 + col] = "Разные";
                        col++;
                    }
                    if (checkLimits.ShowHiHi)
                    {
                        worksheet.Cells[row, 6 + col] = "Разные";
                        col++;
                    }
                    myColumn = col;
                }
                row++;
            }

 //           worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[1, 5]].Columns.AutoFit();
 //           worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 6+ myColumn]].Columns.AutoFit();
            worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[worksheet.UsedRange.Rows.Count, 5]].Columns.AutoFit();
            //worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 6 + myColumn]].Columns.AutoFit();
            //worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 6 + myColumn]].Columns.ColumnWidth;
            worksheet.Range[worksheet.Cells[1, 6 + myColumn - 1], worksheet.Cells[worksheet.UsedRange.Rows.Count, 6 + myColumn - 1]].Columns.AutoFit();
            //worksheet.Range[worksheet.Cells[1, 6], worksheet.Cells[worksheet.UsedRange.Rows.Count, 6 + myColumn-1]].Columns.AutoFit();
            //worksheet.Range[worksheet.Cells[1, 6], worksheet.Cells[1, 6 + myColumn - 2]].Columns.AutoFit();

            //worksheet.Cells[1, worksheet.UsedRange.Columns.Count].Columns.AutoFit();
            //worksheet.Range[worksheet.Cells[3, 4], worksheet.Cells[worksheet.UsedRange.Rows.Count, worksheet.UsedRange.Columns.Count]].Columns.AutoFit();

            excelApplication.ScreenUpdating = true;
            excelApplication.DisplayAlerts = true;

        }

        public bool VarSpecsCommonLimitsCheckSheet()
        {
            return (worksheet.Cells[1, 1].Value == "Отметка" && worksheet.Cells[1, 2].Value == "Оборудование" && worksheet.Cells[1, 3].Value == "Путь" && worksheet.Cells[1, 4].Value == "Имя" && worksheet.Cells[1, 5].Value == "Тег" && worksheet.Cells[1, 6].Value == "Все продукты");
        }

        /// <summary>
        /// Сохранение общих выбранных границ для выбранных тегов
        /// </summary>
        public void VarSpecsCommonLimitSaveFromSheet()
        {
            dynamic dynProductionUnit_Desc;
            dynamic dynTag;
            dynamic dynProductCode;
            dynamic dynPathAF;
            dynamic dynName;


            dynamic dynLoLo;
            dynamic dynLo;
            dynamic dynTarget;
            dynamic dynHiHi;
            dynamic dynHi;
            dynamic dynAuthor;

            VarSpecsChangeStatus resultquery;

            int myrows = worksheet.UsedRange.Rows.Count;
            int mycolumns = worksheet.UsedRange.Columns.Count;

            if (VarSpecsCommonLimitsCheckSheet())
            {
                //Определяем автора изменяющего значения границ
                string author = string.Format(Environment.UserDomainName + "/" + Environment.UserName);


                // Количество границ которые нужно записать
                int SpecsLimits = worksheet.Cells[1, 6].MergeArea.Columns.Count;

                string ProductionUnit_Desc;
                string Tag;
                string ProductCode;
                List<ProductLimit> listlimit = new List<ProductLimit>();


                bool WriteTOSpecsFlag;
                //    //С тройки начинается потому-что заголовок таблицы занимает 2 строки
                for (int rowcount = 3; rowcount < myrows + 1; rowcount++)
                {
                    WriteTOSpecsFlag = true;
                    if (worksheet.Cells[rowcount, 1].Value == "x")
                    {
                        //Получаем значения ProductionUnit_Desc, Tag
                        dynProductionUnit_Desc = worksheet.Cells[rowcount, 2].Value;
                        dynTag = worksheet.Cells[rowcount, 5].Value;

                        ProductionUnit_Desc = (string)dynProductionUnit_Desc;
                        Tag = (string)dynTag;

                        //ProductCode = null потомучто нам не важн код продукта
                        ProductLimit prlimit = new ProductLimit(ProductionUnit_Desc, Tag, null, author);

                        
                        //Здесь пробегаемся по границам продуктов для каждого тега
                            for (int limitsCount = 0; limitsCount < SpecsLimits; limitsCount++)
                            {

                                if ((string)worksheet.Cells[2, 6 + limitsCount].Value == "LoLo")
                                {
                                    dynLoLo = worksheet.Cells[rowcount, 6 + limitsCount].Value;

                                    if ((dynLoLo is double) || (dynLoLo == null))
                                    {
                                        if (dynLoLo == null)
                                        {
                                            prlimit.LoLo = null;
                                        }
                                        else
                                        {
                                            prlimit.LoLo = (decimal)dynLoLo;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        WriteTOSpecsFlag = WriteTOSpecsFlag & false;
                                    }
                                    //if ((dynLoLo is double) && (dynproduct_Code is string))
                                    //if ((dynproduct_Id == null) && (dynproduct_Code is string))

                                    // для тестовой записи ниже табл со спецификациями 
                                    //worksheet.Cells[rowcount+ myrows+3, columncount + limitsCount].Value=dynLoLo;
                                }
                                if ((string)worksheet.Cells[2, 6 + limitsCount].Value == "Lo")
                                {
                                    dynLo = worksheet.Cells[rowcount, 6 + limitsCount].Value;
                                    if ((dynLo is double) || (dynLo == null))
                                    {
                                        if (dynLo == null)
                                        {
                                            prlimit.Lo = null;
                                        }
                                        else
                                        {
                                            prlimit.Lo = (decimal)dynLo;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        WriteTOSpecsFlag = WriteTOSpecsFlag & false;
                                    }

                                // для тестовой записи ниже табл со спецификациями 
                                //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value =dynLo;
                            }
                            if ((string)worksheet.Cells[2, 6 + limitsCount].Value == "Target")
                                {
                                    dynTarget = worksheet.Cells[rowcount, 6 + limitsCount].Value;
                                    if ((dynTarget is double) || (dynTarget == null))
                                    {
                                        if (dynTarget == null)
                                        {
                                            prlimit.Target = null;
                                        }
                                        else
                                        {
                                            prlimit.Target = (decimal)dynTarget;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        WriteTOSpecsFlag = WriteTOSpecsFlag & false;
                                    }

                                // для тестовой записи ниже табл со спецификациями 
                                //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value =dynTarget;
                            }
                            if ((string)worksheet.Cells[2, 6 + limitsCount].Value == "HiHi")
                                {
                                    dynHiHi = worksheet.Cells[rowcount, 6 + limitsCount].Value;
                                    if ((dynHiHi is double) || (dynHiHi == null))
                                    {
                                        if (dynHiHi == null)
                                        {
                                            prlimit.HiHi = null;
                                        }
                                        else
                                        {
                                            prlimit.HiHi = (decimal)dynHiHi;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        WriteTOSpecsFlag = WriteTOSpecsFlag & false;
                                    }
                                // для тестовой записи ниже табл со спецификациями 
                                //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value = dynHiHi;
                            }
                            if ((string)worksheet.Cells[2, 6 + limitsCount].Value == "Hi")
                                {
                                    dynHi = worksheet.Cells[rowcount, 6 + limitsCount].Value;

                                    if ((dynHi is double) || (dynHi == null))
                                    {
                                        if (dynHi == null)
                                        {
                                            prlimit.Hi = null;
                                        }
                                        else
                                        {
                                            prlimit.Hi = (decimal)dynHi;
                                        }
                                        continue;
                                    }
                                    else
                                    {
                                        WriteTOSpecsFlag = WriteTOSpecsFlag & false;
                                    }
                                // для тестовой записи ниже табл со спецификациями 
                                //worksheet.Cells[rowcount + myrows + 3, columncount + limitsCount].Value = dynHi;
                            }
                        }
                        if (WriteTOSpecsFlag)
                        {
                            resultquery = specs.UpDateVarSpecsCommonBound(prlimit);
                            // !!!!! Здесь нужно записать Отдельный экземпляр ProductLimits В БД Specs
                            // listlimit.Add(prlimit);
                            //resultquery = specs.UpDateVarSpecs(prlimit);

                            if (resultquery == VarSpecsChangeStatus.ChangeOK)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[rowcount, 6], worksheet.Cells[rowcount, 6 + SpecsLimits - 1]], true);
                            }
                            else if (resultquery == VarSpecsChangeStatus.ChangeBad)
                            {
                                MarkRange(worksheet.Range[worksheet.Cells[rowcount, 6], worksheet.Cells[rowcount, 6 + SpecsLimits - 1]], false);
                            }
                        }

                    }
                }

            }
            else
            {
                //Данные на листе не соответствуют выбранной операции
                MessageBox.Show("Данные на листе не соответствуют выбранной операции", "Сообщение об ошибке", MessageBoxButtons.OK);
            }
        }

        public void TempMethod(string unit)
        {
            List<EntityVarSpecCommonBound> list = specs.GetCommonBound(unit);
            int i = list.Count;
        }
    }
}
