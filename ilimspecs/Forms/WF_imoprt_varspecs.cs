using ilimspecs.BL;
using ilimspecs.Data;
using Microsoft.Office.Interop.Excel;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ilimspecs.Forms
{
    public partial class WF_imoprt_varspecs : Form
    {
        //для всплывающих подсказок 
        ToolTip toolTipVarSpecs = new ToolTip();
        int ttIndexVarSpecs = -1;

        ToolTip toolTipProduct = new ToolTip();
        int ttIndexProduct = -1;


        public WF_imoprt_varspecs()
        {
            InitializeComponent();
        }


        private void WF_imoprt_varspecs_Load(object sender, EventArgs e)
        {
            
            //AFEnvironment af = new AFEnvironment();
            //List<AFElement> aflist = af.GetProductUnitFromAF();

            AFEnvironment afe = new AFEnvironment();
            AFDatabase afdb = afe.GetDatabase();

            afTreeViewContolVarSpecs.ShowLayers = false;
            afTreeViewContolVarSpecs.SetAFRoot(afdb.Elements["Контроль"], null, "");

            afTreeViewContolVarSpecs.BeginUpdate();
            afTreeViewContolVarSpecs.ExpandAll();
            afTreeViewContolVarSpecs.CollapseAll();
            afTreeViewContolVarSpecs.EndUpdate();

        }

        private void afTreeViewContolVarSpecs_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //lstViewProduct.Items.Clear();
            //lstViewVarSpec.Items.Clear();
            chcklstProduct.Items.Clear();
            chcklstVarSpecs.Items.Clear();


            AFTreeNode aftrn = e.Node as AFTreeNode;
            //AFTreeNode aftrn = afTreeView2.SelectedNode as AFTreeNode;
            if (aftrn != null)
            {
                AFElement afe = aftrn.AFObject as AFElement;
                if (afe != null)
                {
                    //TODO: Переделать на константы которые возвращают "Код производственного участка"
                    AFAttribute afatribute = afe.Attributes["Код производственного участка"];
                    if (afatribute != null)
                    {
                        string CPS = afatribute.GetValue().ToString();
                        Specs specs = new Specs();
                        AFEnvironment afenv = new AFEnvironment();

                        List<EntityVarSpecsAF> lstspecification = afenv.GetSpecificationAF(afe);
                        List<EntityProduct> lstproduct = specs.GetAllProductByUnit(CPS);

                        if (lstproduct != null)
                        {
                            chcklstProduct.BeginUpdate();//отключаю перерисовку для вывода элементов
                            chcklstProduct.Items.AddRange(lstproduct.ToArray<EntityProduct>());
                            //Выбираем все элементы
                            for (int i = 0; i < chcklstProduct.Items.Count; i++)
                            {
                                chcklstProduct.SetItemChecked(i, true);
                            }

                            chcklstProduct.EndUpdate();
                        }
                        if (lstspecification != null)
                        {
                            chcklstProduct.BeginUpdate();//отключаю перерисовку для вывода элементов
                            chcklstVarSpecs.Items.AddRange(lstspecification.ToArray<EntityVarSpecsAF>());

                            //Выбираем все элементы
                            for (int i = 0; i < chcklstVarSpecs.Items.Count; i++)
                            {
                                chcklstVarSpecs.SetItemChecked(i, true);
                            }

                            chcklstProduct.EndUpdate();
                        }
                    
                    }
                }
            }

        }


        private void chcklstVarSpecs_MouseHover(object sender, EventArgs e)
        {
            ShowToolTip();
        }

        private void chcklstVarSpecs_MouseMove(object sender, MouseEventArgs e)
        {
            if (ttIndexVarSpecs != chcklstVarSpecs.IndexFromPoint(e.Location))
                ShowToolTip();
        }

        private void ShowToolTip()
        {
            ttIndexVarSpecs = chcklstVarSpecs.IndexFromPoint(chcklstVarSpecs.PointToClient(MousePosition));
            if (ttIndexVarSpecs > -1)
            {
                System.Drawing.Point p = PointToClient(MousePosition);
                toolTipVarSpecs.ToolTipTitle = "Имя Тега";

                EntityVarSpecsAF varspecs = (EntityVarSpecsAF)chcklstVarSpecs.Items[ttIndexVarSpecs];
                toolTipVarSpecs.SetToolTip(chcklstVarSpecs, varspecs.Name);
            }
        }

        private void chcklstProduct_MouseHover(object sender, EventArgs e)
        {
            //ttIndexProduct
            ShowToolTipProduct();
        }

        private void chcklstProduct_MouseMove(object sender, MouseEventArgs e)
        {
            if (ttIndexProduct != chcklstProduct.IndexFromPoint(e.Location))
                ShowToolTipProduct();

        }

        private void ShowToolTipProduct()
        {
            ttIndexProduct = chcklstProduct.IndexFromPoint(chcklstProduct.PointToClient(MousePosition));
            if (ttIndexProduct > -1)
            {
                System.Drawing.Point p = PointToClient(MousePosition);
                toolTipProduct.ToolTipTitle = "Код продукта";

                EntityProduct product = (EntityProduct)chcklstProduct.Items[ttIndexProduct];
                toolTipProduct.SetToolTip(chcklstProduct, product.Product_Code);
            }
        }

        private void btnVarSpecCheckAll_Click(object sender, EventArgs e)
        {
            //bool flag = e.NewValue == CheckState.Checked ? true : false;
            for (int i = 0; i < chcklstVarSpecs.Items.Count; i++)
                chcklstVarSpecs.SetItemChecked(i, true);
        }

        private void btnVarSpecNotCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chcklstVarSpecs.Items.Count; i++)
                chcklstVarSpecs.SetItemChecked(i, false);
        }

        private void btnProductCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chcklstProduct.Items.Count; i++)
                chcklstProduct.SetItemChecked(i, true);
        }

        private void btnProductNotCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chcklstProduct.Items.Count; i++)
                chcklstProduct.SetItemChecked(i, false);
        }

        private void btnLoadAllLimits_Click(object sender, EventArgs e)
        {
            //List<EntityVarSpecsAF> listFromCheckBoxVarSpec=new List<EntityVarSpecsAF>();

            //foreach (var item in chcklstVarSpecs.CheckedItems)
            //{
            //    listFromCheckBoxVarSpec.Add((EntityVarSpecsAF)item);
            //}

            //List<EntitySpecification> listvarspec = specs.GetAllSpecification(listFromCheckBoxVarSpec);

            //this.Close();

            if (chcklstVarSpecs.CheckedItems.Count == 0)
            {
                MessageBox.Show("Отсутствуют выбранные спецификации", "Сообщение об ошибке", MessageBoxButtons.OK);
                return;
            }

            List<EntityVarSpecsAF> listFromCheckBoxVarSpec = new List<EntityVarSpecsAF>();

            foreach (var item in chcklstVarSpecs.CheckedItems)
            {
                listFromCheckBoxVarSpec.Add((EntityVarSpecsAF)item);
            }

            CheckLimits checkLimits = new CheckLimits();
            checkLimits.ShowLoLo = chkLoLo.Checked;
            checkLimits.ShowLo = chkLo.Checked;
            checkLimits.ShowTarget = chkTarget.Checked;
            checkLimits.ShowHi = chkHi.Checked;
            checkLimits.ShowHiHi = chkHiHi.Checked;
            checkLimits.ShowAuthor = chkAuthor.Checked;


            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);
            excelBook.VarSpecsCommonLimitsTablePrint(listFromCheckBoxVarSpec,  checkLimits);

            this.Close();



        }

        private void btnLoadAllProduct_Click(object sender, EventArgs e)
        {
            if (chcklstProduct.CheckedItems.Count==0)
            {
                MessageBox.Show("Отсутствуют выбранные продукты", "Сообщение об ошибке", MessageBoxButtons.OK);
                return;
            }

            if (chcklstVarSpecs.CheckedItems.Count == 0)
            {
                MessageBox.Show("Отсутствуют выбранные спецификации", "Сообщение об ошибке", MessageBoxButtons.OK);
                return;
            }

            List<EntityVarSpecsAF> listFromCheckBoxVarSpec = new List<EntityVarSpecsAF>();

            foreach (var item in chcklstVarSpecs.CheckedItems)
            {
                listFromCheckBoxVarSpec.Add((EntityVarSpecsAF)item);
            }

            CheckLimits checkLimits = new CheckLimits();
            checkLimits.ShowLoLo = chkLoLo.Checked;
            checkLimits.ShowLo = chkLo.Checked;
            checkLimits.ShowTarget = chkTarget.Checked;
            checkLimits.ShowHi = chkHi.Checked;
            checkLimits.ShowHiHi = chkHiHi.Checked;
            checkLimits.ShowAuthor = chkAuthor.Checked;

            List<string> prcode = new List<string>();

            foreach (var item in chcklstProduct.CheckedItems)
            {
                prcode.Add(((EntityProduct)item).Product_Code);
            }

            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);
            excelBook.VarSpecsTablePrint(listFromCheckBoxVarSpec, prcode, checkLimits);
           
            this.Close();
            
        }
    }
}
