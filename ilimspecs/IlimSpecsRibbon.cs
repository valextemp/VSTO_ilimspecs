using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ilimspecs.BL;
using Microsoft.Office.Interop.Excel;
using ilimspecs.Forms;
using System.Windows.Forms;

namespace ilimspecs
{
    public partial class IlimSpecsRibbon
    {
        private void IlimSpecsRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnImportProductionUnit_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);

            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            try
            {
                excelBook.ProductionUnitTablePrint();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                worksheet.Application.Cursor = initialCursor;
            }
        }

        private void btnExportProductionUnit_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);

            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            try
            {
                excelBook.ProductionUnitSaveFromSheet();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                worksheet.Application.Cursor = initialCursor;
                MessageBox.Show("Запись данных в базу данных завершена", "Информационное сообщение", MessageBoxButtons.OK);
            }

        }

        private void btnImportProduct_Click(object sender, RibbonControlEventArgs e)
        {
            WF_import_products improdform = new WF_import_products();
            improdform.ShowDialog();
        }

        private void btnImportVarSpec_Click(object sender, RibbonControlEventArgs e)
        {
            WF_imoprt_varspecs improdform = new WF_imoprt_varspecs();
            improdform.ShowDialog();

        }

        private void btnExprotProduct_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);

            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            try
            {
                excelBook.ProductSaveFromSheet();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                worksheet.Application.Cursor = initialCursor;
                MessageBox.Show("Запись данных в базу данных завершена", "Информационное сообщение", MessageBoxButtons.OK);
            }



        }


        private void btnExportVarSpec_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);

            var initialCursor = excelApplication.Cursor; //Сохраняем предидущий курсор
            worksheet.Application.Cursor = XlMousePointer.xlWait;

            try
            {
                if(!excelBook.VarSpecsCommonLimitsCheckSheet())
                {
                    excelBook.VarSpecsSaveFromSheet();
                }
                else
                {
                    excelBook.VarSpecsCommonLimitSaveFromSheet();
                }
                
            }
            catch (Exception)
            {
            }
            finally
            {
                worksheet.Application.Cursor = initialCursor;
                MessageBox.Show("Запись данных в базу данных завершена", "Информационное сообщение", MessageBoxButtons.OK);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            
            //ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);
            //excelBook.TempMethod("OMC");
        }
    }
}
