using ilimspecs.Data;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using ilimspecs.BL;

namespace ilimspecs.Forms
{
    public partial class WF_import_products : Form
    {

        public WF_import_products()
        {
            InitializeComponent();
        }

        private void WF_import_products_Load(object sender, EventArgs e)
        {
            AFEnvironment afe = new AFEnvironment();
            AFDatabase afdb = afe.GetDatabase();

            afTreeViewControl.ShowLayers = false;
            afTreeViewControl.SetAFRoot(afdb.Elements["Контроль"], null, "");

            afTreeViewControl.BeginUpdate();
            afTreeViewControl.ExpandAll();
            afTreeViewControl.CollapseAll();
            afTreeViewControl.EndUpdate();
            
            List<AFElement> lae = afe.GetProductUnitFromAF();

            TreeNodeCollection nodes = afTreeViewControl.Nodes;

            afTreeViewControl.BeginUpdate();
            int iii = 0;

            while (nodes.Count > iii)
            {
                DeleteAFElement(nodes[iii], lae);
                iii++;
            }
            afTreeViewControl.EndUpdate();
            afTreeViewControl.Refresh();

           
        }

        private void DeleteAFElement(TreeNode treeNode, List<AFElement> listAFE)
        {

            AFTreeNode aftrn = treeNode as AFTreeNode;
            if (aftrn != null)
            {
                AFElement afe = aftrn.AFObject as AFElement;
                if (afe != null)
                {
                    foreach (AFElement item in listAFE)
                    {
                        if (item.Equals(afe))
                        {
                            treeNode.Nodes.Clear();
                        }
                    }
                }
            }

            int ii = 0;
            while (treeNode.Nodes.Count > ii)
            {
                DeleteAFElement(treeNode.Nodes[ii], listAFE);
                ii++;
            }

        }

        private void afTreeViewControl_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            lstProducts.Items.Clear();

            AFTreeNode aftrn = e.Node as AFTreeNode;
            //AFTreeNode aftrn = afTreeView2.SelectedNode as AFTreeNode;
            if (aftrn != null)
            {
                AFElement afe = aftrn.AFObject as AFElement;
                if (afe != null)
                {
                    AFAttribute afatribute = afe.Attributes["Код производственного участка"];
                    if (afatribute != null)
                    {
                        string CPS = afatribute.GetValue().ToString();
                        Specs specs = new Specs();
                        List<EntityProduct> lstproduct = specs.GetAllProductByUnit(CPS);
                        if (lstproduct != null)
                        {
                            lstProducts.Items.AddRange(lstproduct.ToArray<EntityProduct>());
                        }
                    }
                }
            }

        }

        private void btnLoadProduct_Click(object sender, EventArgs e)
        {
            if (lstProducts.Items==null)
            {
                MessageBox.Show("Отсутсвтвуют выбранные продукты", "Сообщение об ошибке", MessageBoxButtons.OK);
                return;
            }

            Microsoft.Office.Interop.Excel.Application excelApplication = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;


            ExcelBook excelBook = new ExcelBook(excelApplication, workbook, worksheet);
            excelBook.ProductTablePrint(lstProducts.Items.Cast<EntityProduct>().ToList());
            this.Close();

        }
    }
}
