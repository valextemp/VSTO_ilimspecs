namespace ilimspecs.Forms
{
    partial class WF_imoprt_varspecs
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pnlAFControl = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.afTreeViewContolVarSpecs = new OSIsoft.AF.UI.AFTreeView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chcklstVarSpecs = new System.Windows.Forms.CheckedListBox();
            this.chcklstProduct = new System.Windows.Forms.CheckedListBox();
            this.btnProductNotCheckAll = new System.Windows.Forms.Button();
            this.btnProductCheckAll = new System.Windows.Forms.Button();
            this.btnVarSpecNotCheckAll = new System.Windows.Forms.Button();
            this.btnVarSpecCheckAll = new System.Windows.Forms.Button();
            this.lblProduct = new System.Windows.Forms.Label();
            this.lblVarSpec = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkAuthor = new System.Windows.Forms.CheckBox();
            this.chkHiHi = new System.Windows.Forms.CheckBox();
            this.chkHi = new System.Windows.Forms.CheckBox();
            this.chkTarget = new System.Windows.Forms.CheckBox();
            this.chkLo = new System.Windows.Forms.CheckBox();
            this.chkLoLo = new System.Windows.Forms.CheckBox();
            this.btnLoadAllProduct = new System.Windows.Forms.Button();
            this.btnLoadAllLimits = new System.Windows.Forms.Button();
            this.pnlAFControl.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlAFControl
            // 
            this.pnlAFControl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlAFControl.Controls.Add(this.label1);
            this.pnlAFControl.Controls.Add(this.afTreeViewContolVarSpecs);
            this.pnlAFControl.Location = new System.Drawing.Point(2, 6);
            this.pnlAFControl.Name = "pnlAFControl";
            this.pnlAFControl.Size = new System.Drawing.Size(330, 465);
            this.pnlAFControl.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "AF Контроль";
            // 
            // afTreeViewContolVarSpecs
            // 
            this.afTreeViewContolVarSpecs.HideSelection = false;
            this.afTreeViewContolVarSpecs.Location = new System.Drawing.Point(3, 25);
            this.afTreeViewContolVarSpecs.Name = "afTreeViewContolVarSpecs";
            this.afTreeViewContolVarSpecs.ShowNodeToolTips = true;
            this.afTreeViewContolVarSpecs.Size = new System.Drawing.Size(322, 432);
            this.afTreeViewContolVarSpecs.TabIndex = 2;
            this.afTreeViewContolVarSpecs.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.afTreeViewContolVarSpecs_NodeMouseClick);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.chcklstVarSpecs);
            this.panel1.Controls.Add(this.chcklstProduct);
            this.panel1.Controls.Add(this.btnProductNotCheckAll);
            this.panel1.Controls.Add(this.btnProductCheckAll);
            this.panel1.Controls.Add(this.btnVarSpecNotCheckAll);
            this.panel1.Controls.Add(this.btnVarSpecCheckAll);
            this.panel1.Controls.Add(this.lblProduct);
            this.panel1.Controls.Add(this.lblVarSpec);
            this.panel1.Location = new System.Drawing.Point(338, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(573, 465);
            this.panel1.TabIndex = 1;
            // 
            // chcklstVarSpecs
            // 
            this.chcklstVarSpecs.CheckOnClick = true;
            this.chcklstVarSpecs.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chcklstVarSpecs.FormattingEnabled = true;
            this.chcklstVarSpecs.HorizontalScrollbar = true;
            this.chcklstVarSpecs.Location = new System.Drawing.Point(4, 26);
            this.chcklstVarSpecs.Name = "chcklstVarSpecs";
            this.chcklstVarSpecs.Size = new System.Drawing.Size(289, 394);
            this.chcklstVarSpecs.TabIndex = 9;
            this.chcklstVarSpecs.MouseHover += new System.EventHandler(this.chcklstVarSpecs_MouseHover);
            this.chcklstVarSpecs.MouseMove += new System.Windows.Forms.MouseEventHandler(this.chcklstVarSpecs_MouseMove);
            // 
            // chcklstProduct
            // 
            this.chcklstProduct.CheckOnClick = true;
            this.chcklstProduct.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.chcklstProduct.FormattingEnabled = true;
            this.chcklstProduct.HorizontalScrollbar = true;
            this.chcklstProduct.Location = new System.Drawing.Point(299, 26);
            this.chcklstProduct.Name = "chcklstProduct";
            this.chcklstProduct.Size = new System.Drawing.Size(255, 394);
            this.chcklstProduct.TabIndex = 8;
            this.chcklstProduct.ThreeDCheckBoxes = true;
            this.chcklstProduct.MouseHover += new System.EventHandler(this.chcklstProduct_MouseHover);
            this.chcklstProduct.MouseMove += new System.Windows.Forms.MouseEventHandler(this.chcklstProduct_MouseMove);
            // 
            // btnProductNotCheckAll
            // 
            this.btnProductNotCheckAll.Location = new System.Drawing.Point(443, 434);
            this.btnProductNotCheckAll.Name = "btnProductNotCheckAll";
            this.btnProductNotCheckAll.Size = new System.Drawing.Size(101, 23);
            this.btnProductNotCheckAll.TabIndex = 7;
            this.btnProductNotCheckAll.Text = "Все не выбрано";
            this.btnProductNotCheckAll.UseVisualStyleBackColor = true;
            this.btnProductNotCheckAll.Click += new System.EventHandler(this.btnProductNotCheckAll_Click);
            // 
            // btnProductCheckAll
            // 
            this.btnProductCheckAll.Location = new System.Drawing.Point(325, 434);
            this.btnProductCheckAll.Name = "btnProductCheckAll";
            this.btnProductCheckAll.Size = new System.Drawing.Size(101, 23);
            this.btnProductCheckAll.TabIndex = 6;
            this.btnProductCheckAll.Text = "Выбрать все";
            this.btnProductCheckAll.UseVisualStyleBackColor = true;
            this.btnProductCheckAll.Click += new System.EventHandler(this.btnProductCheckAll_Click);
            // 
            // btnVarSpecNotCheckAll
            // 
            this.btnVarSpecNotCheckAll.Location = new System.Drawing.Point(164, 434);
            this.btnVarSpecNotCheckAll.Name = "btnVarSpecNotCheckAll";
            this.btnVarSpecNotCheckAll.Size = new System.Drawing.Size(101, 23);
            this.btnVarSpecNotCheckAll.TabIndex = 5;
            this.btnVarSpecNotCheckAll.Text = "Все не выбрано";
            this.btnVarSpecNotCheckAll.UseVisualStyleBackColor = true;
            this.btnVarSpecNotCheckAll.Click += new System.EventHandler(this.btnVarSpecNotCheckAll_Click);
            // 
            // btnVarSpecCheckAll
            // 
            this.btnVarSpecCheckAll.Location = new System.Drawing.Point(46, 434);
            this.btnVarSpecCheckAll.Name = "btnVarSpecCheckAll";
            this.btnVarSpecCheckAll.Size = new System.Drawing.Size(101, 23);
            this.btnVarSpecCheckAll.TabIndex = 4;
            this.btnVarSpecCheckAll.Text = "Выбрать все";
            this.btnVarSpecCheckAll.UseVisualStyleBackColor = true;
            this.btnVarSpecCheckAll.Click += new System.EventHandler(this.btnVarSpecCheckAll_Click);
            // 
            // lblProduct
            // 
            this.lblProduct.AutoSize = true;
            this.lblProduct.Location = new System.Drawing.Point(305, 6);
            this.lblProduct.Name = "lblProduct";
            this.lblProduct.Size = new System.Drawing.Size(57, 13);
            this.lblProduct.TabIndex = 3;
            this.lblProduct.Text = "Продукты";
            // 
            // lblVarSpec
            // 
            this.lblVarSpec.AutoSize = true;
            this.lblVarSpec.Location = new System.Drawing.Point(7, 6);
            this.lblVarSpec.Name = "lblVarSpec";
            this.lblVarSpec.Size = new System.Drawing.Size(82, 13);
            this.lblVarSpec.TabIndex = 1;
            this.lblVarSpec.Text = "Спецификации";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkAuthor);
            this.groupBox1.Controls.Add(this.chkHiHi);
            this.groupBox1.Controls.Add(this.chkHi);
            this.groupBox1.Controls.Add(this.chkTarget);
            this.groupBox1.Controls.Add(this.chkLo);
            this.groupBox1.Controls.Add(this.chkLoLo);
            this.groupBox1.Location = new System.Drawing.Point(917, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(114, 179);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Параметры";
            // 
            // chkAuthor
            // 
            this.chkAuthor.AutoSize = true;
            this.chkAuthor.Location = new System.Drawing.Point(16, 147);
            this.chkAuthor.Name = "chkAuthor";
            this.chkAuthor.Size = new System.Drawing.Size(56, 17);
            this.chkAuthor.TabIndex = 5;
            this.chkAuthor.Text = "Автор";
            this.chkAuthor.UseVisualStyleBackColor = true;
            // 
            // chkHiHi
            // 
            this.chkHiHi.AutoSize = true;
            this.chkHiHi.Checked = true;
            this.chkHiHi.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHiHi.Location = new System.Drawing.Point(16, 123);
            this.chkHiHi.Name = "chkHiHi";
            this.chkHiHi.Size = new System.Drawing.Size(46, 17);
            this.chkHiHi.TabIndex = 4;
            this.chkHiHi.Text = "HiHi";
            this.chkHiHi.UseVisualStyleBackColor = true;
            // 
            // chkHi
            // 
            this.chkHi.AutoSize = true;
            this.chkHi.Checked = true;
            this.chkHi.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHi.Location = new System.Drawing.Point(16, 99);
            this.chkHi.Name = "chkHi";
            this.chkHi.Size = new System.Drawing.Size(36, 17);
            this.chkHi.TabIndex = 3;
            this.chkHi.Text = "Hi";
            this.chkHi.UseVisualStyleBackColor = true;
            // 
            // chkTarget
            // 
            this.chkTarget.AutoSize = true;
            this.chkTarget.Checked = true;
            this.chkTarget.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkTarget.Location = new System.Drawing.Point(16, 75);
            this.chkTarget.Name = "chkTarget";
            this.chkTarget.Size = new System.Drawing.Size(57, 17);
            this.chkTarget.TabIndex = 2;
            this.chkTarget.Text = "Target";
            this.chkTarget.UseVisualStyleBackColor = true;
            // 
            // chkLo
            // 
            this.chkLo.AutoSize = true;
            this.chkLo.Checked = true;
            this.chkLo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLo.Location = new System.Drawing.Point(16, 51);
            this.chkLo.Name = "chkLo";
            this.chkLo.Size = new System.Drawing.Size(38, 17);
            this.chkLo.TabIndex = 1;
            this.chkLo.Text = "Lo";
            this.chkLo.UseVisualStyleBackColor = true;
            // 
            // chkLoLo
            // 
            this.chkLoLo.AutoSize = true;
            this.chkLoLo.Checked = true;
            this.chkLoLo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLoLo.Location = new System.Drawing.Point(16, 27);
            this.chkLoLo.Name = "chkLoLo";
            this.chkLoLo.Size = new System.Drawing.Size(50, 17);
            this.chkLoLo.TabIndex = 0;
            this.chkLoLo.Text = "LoLo";
            this.chkLoLo.UseVisualStyleBackColor = true;
            // 
            // btnLoadAllProduct
            // 
            this.btnLoadAllProduct.Location = new System.Drawing.Point(922, 353);
            this.btnLoadAllProduct.Name = "btnLoadAllProduct";
            this.btnLoadAllProduct.Size = new System.Drawing.Size(109, 49);
            this.btnLoadAllProduct.TabIndex = 5;
            this.btnLoadAllProduct.Text = "Выгрузить на лист выбранные продукты";
            this.btnLoadAllProduct.UseVisualStyleBackColor = true;
            this.btnLoadAllProduct.Click += new System.EventHandler(this.btnLoadAllProduct_Click);
            // 
            // btnLoadAllLimits
            // 
            this.btnLoadAllLimits.Location = new System.Drawing.Point(922, 408);
            this.btnLoadAllLimits.Name = "btnLoadAllLimits";
            this.btnLoadAllLimits.Size = new System.Drawing.Size(109, 63);
            this.btnLoadAllLimits.TabIndex = 6;
            this.btnLoadAllLimits.Text = "Выгрузить на лист общие лимиты для всех продуктов";
            this.btnLoadAllLimits.UseVisualStyleBackColor = true;
            this.btnLoadAllLimits.Click += new System.EventHandler(this.btnLoadAllLimits_Click);
            // 
            // WF_imoprt_varspecs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1038, 476);
            this.Controls.Add(this.btnLoadAllLimits);
            this.Controls.Add(this.btnLoadAllProduct);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.pnlAFControl);
            this.Name = "WF_imoprt_varspecs";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Импорт спецификаций";
            this.Load += new System.EventHandler(this.WF_imoprt_varspecs_Load);
            this.pnlAFControl.ResumeLayout(false);
            this.pnlAFControl.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlAFControl;
        private System.Windows.Forms.Label label1;
        private OSIsoft.AF.UI.AFTreeView afTreeViewContolVarSpecs;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnProductNotCheckAll;
        private System.Windows.Forms.Button btnProductCheckAll;
        private System.Windows.Forms.Button btnVarSpecNotCheckAll;
        private System.Windows.Forms.Button btnVarSpecCheckAll;
        private System.Windows.Forms.Label lblProduct;
        private System.Windows.Forms.Label lblVarSpec;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkAuthor;
        private System.Windows.Forms.CheckBox chkHiHi;
        private System.Windows.Forms.CheckBox chkHi;
        private System.Windows.Forms.CheckBox chkTarget;
        private System.Windows.Forms.CheckBox chkLo;
        private System.Windows.Forms.CheckBox chkLoLo;
        private System.Windows.Forms.CheckedListBox chcklstProduct;
        private System.Windows.Forms.CheckedListBox chcklstVarSpecs;
        private System.Windows.Forms.Button btnLoadAllProduct;
        private System.Windows.Forms.Button btnLoadAllLimits;
    }
}