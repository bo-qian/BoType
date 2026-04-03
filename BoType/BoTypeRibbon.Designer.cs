namespace BoType
{
    partial class BoTypeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public BoTypeRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.comboBoxWidth = this.Factory.CreateRibbonComboBox();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.boxRef = this.Factory.CreateRibbonBox();
            this.dropDownRefEq = this.Factory.CreateRibbonDropDown();
            this.dropDownRefStyle = this.Factory.CreateRibbonDropDown();
            this.buttonSetDefaultRefStyle = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.groupAbout = this.Factory.CreateRibbonGroup();
            this.buttonGithub = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.box1.SuspendLayout();
            this.group3.SuspendLayout();
            this.boxRef.SuspendLayout();
            this.groupAbout.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.groupAbout);
            this.tab1.Label = "BoType";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "插入公式";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "单行公式";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "EquationInsertNew";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Label = "行内公式";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "EquationNormalText";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.box1);
            this.group2.Items.Add(this.button5);
            this.group2.Label = "公式编号";
            this.group2.Name = "group2";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.dropDown1);
            this.box1.Items.Add(this.comboBoxWidth);
            this.box1.Items.Add(this.button6);
            this.box1.Name = "box1";
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "无";
            ribbonDropDownItemImpl2.Label = "(1)";
            ribbonDropDownItemImpl3.Label = "(1.1)";
            ribbonDropDownItemImpl4.Label = "(1-1)";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown1.Label = "编号样式:";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.OfficeImageId = "NumberingGallery";
            this.dropDown1.ShowImage = true;
            this.dropDown1.SizeString = "MMMMMMIII";
            // 
            // comboBoxWidth
            // 
            ribbonDropDownItemImpl5.Label = "28 磅";
            ribbonDropDownItemImpl6.Label = "30 磅";
            ribbonDropDownItemImpl7.Label = "36 磅";
            ribbonDropDownItemImpl8.Label = "38 磅";
            ribbonDropDownItemImpl9.Label = "40 磅";
            ribbonDropDownItemImpl10.Label = "42 磅";
            ribbonDropDownItemImpl11.Label = "48 磅";
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl5);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl6);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl7);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl8);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl9);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl10);
            this.comboBoxWidth.Items.Add(ribbonDropDownItemImpl11);
            this.comboBoxWidth.Label = "编号占位:";
            this.comboBoxWidth.Name = "comboBoxWidth";
            this.comboBoxWidth.OfficeImageId = "SizeToControlWidth";
            this.comboBoxWidth.ShowImage = true;
            this.comboBoxWidth.SizeString = "MMMMMMM";
            this.comboBoxWidth.Text = "38 磅";
            // 
            // button6
            // 
            this.button6.Label = "设置为默认编号样式";
            this.button6.Name = "button6";
            this.button6.OfficeImageId = "SetAsDefault";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Label = "给公式编号";
            this.button5.Name = "button5";
            this.button5.OfficeImageId = "NumberingRestart";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.boxRef);
            this.group3.Items.Add(this.button3);
            this.group3.Items.Add(this.button4);
            this.group3.Label = "公式引用";
            this.group3.Name = "group3";
            // 
            // boxRef
            // 
            this.boxRef.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.boxRef.Items.Add(this.dropDownRefEq);
            this.boxRef.Items.Add(this.dropDownRefStyle);
            this.boxRef.Items.Add(this.buttonSetDefaultRefStyle);
            this.boxRef.Name = "boxRef";
            // 
            // dropDownRefEq
            // 
            this.dropDownRefEq.Label = "公式列表:";
            this.dropDownRefEq.Name = "dropDownRefEq";
            this.dropDownRefEq.OfficeImageId = "TableOfFiguresInsert";
            this.dropDownRefEq.ShowImage = true;
            this.dropDownRefEq.SizeString = "MMMMMMM";
            // 
            // dropDownRefStyle
            // 
            this.dropDownRefStyle.Label = "引用样式:";
            this.dropDownRefStyle.Name = "dropDownRefStyle";
            this.dropDownRefStyle.OfficeImageId = "CrossReferenceInsert";
            this.dropDownRefStyle.ShowImage = true;
            this.dropDownRefStyle.SizeString = "MMMMMMM";
            // 
            // buttonSetDefaultRefStyle
            // 
            this.buttonSetDefaultRefStyle.Label = "设置为默认引用样式";
            this.buttonSetDefaultRefStyle.Name = "buttonSetDefaultRefStyle";
            this.buttonSetDefaultRefStyle.OfficeImageId = "SetAsDefault";
            this.buttonSetDefaultRefStyle.ShowImage = true;
            this.buttonSetDefaultRefStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSetDefaultRefStyle_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "插入引用";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "CrossReferenceInsert";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Label = "更新引用";
            this.button4.Name = "button4";
            this.button4.OfficeImageId = "Refresh";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // groupAbout
            // 
            this.groupAbout.Items.Add(this.buttonGithub);
            this.groupAbout.Label = "关于";
            this.groupAbout.Name = "groupAbout";
            // 
            // buttonGithub
            // 
            this.buttonGithub.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonGithub.Label = "开源地址";
            this.buttonGithub.Name = "buttonGithub";
            this.buttonGithub.OfficeImageId = "Help";
            this.buttonGithub.ShowImage = true;
            this.buttonGithub.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGithub_Click);
            // 
            // BoTypeRibbon
            // 
            this.Name = "BoTypeRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.BoTypeRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.boxRef.ResumeLayout(false);
            this.boxRef.PerformLayout();
            this.groupAbout.ResumeLayout(false);
            this.groupAbout.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxRef;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownRefEq;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownRefStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSetDefaultRefStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGithub;
    }

    partial class ThisRibbonCollection
    {
        internal BoTypeRibbon BoTypeRibbon
        {
            get { return this.GetRibbon<BoTypeRibbon>(); }
        }
    }
}
