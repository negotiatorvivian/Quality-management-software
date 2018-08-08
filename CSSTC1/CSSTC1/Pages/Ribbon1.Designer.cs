namespace CSSTC1.Pages {
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent() {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.menu3 = this.Factory.CreateRibbonMenu();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.button1 = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog3 = new System.Windows.Forms.OpenFileDialog();
            this.openFileDialog4 = new System.Windows.Forms.OpenFileDialog();
            this.menu4 = this.Factory.CreateRibbonMenu();
            this.button11 = this.Factory.CreateRibbonButton();
            this.menu5 = this.Factory.CreateRibbonMenu();
            this.button12 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.menu3);
            this.group1.Items.Add(this.menu2);
            this.group1.Label = "上传文件";
            this.group1.Name = "group1";
            // 
            // menu3
            // 
            this.menu3.Items.Add(this.button5);
            this.menu3.Items.Add(this.button6);
            this.menu3.Items.Add(this.button7);
            this.menu3.Items.Add(this.button8);
            this.menu3.Label = "填写基本信息";
            this.menu3.Name = "menu3";
            this.menu3.SuperTip = "填写各个阶段的基本信息";
            // 
            // button5
            // 
            this.button5.Label = "项目立项阶段";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Enabled = false;
            this.button6.Label = "测试需求分析与策划";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Enabled = false;
            this.button7.Label = "测试设计与执行";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Enabled = false;
            this.button8.Label = "测试总结";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // menu2
            // 
            this.menu2.Items.Add(this.button1);
            this.menu2.Items.Add(this.menu1);
            this.menu2.Items.Add(this.menu4);
            this.menu2.Items.Add(this.button4);
            this.menu2.Label = "读取文件";
            this.menu2.Name = "menu2";
            this.menu2.SuperTip = "选择要读取的文件，请注意文件内表格排列顺序";
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Label = "项目立项阶段";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // menu1
            // 
            this.menu1.Enabled = false;
            this.menu1.Items.Add(this.button9);
            this.menu1.Items.Add(this.button10);
            this.menu1.Label = "测试需求分析与策划阶段";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // button9
            // 
            this.button9.Label = "测试大纲";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Label = "测试需求规格说明与测试计划";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // button3
            // 
            this.button3.Label = "基本信息表";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Enabled = false;
            this.button4.Label = "测试总结阶段";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Label = "更新";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.Label = "更新文档内容";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "doc";
            this.openFileDialog1.FileName = "质量管理体系文件";
            this.openFileDialog1.Filter = "所有word文档|*.doc; *.docx; *.dot; *.dotx|所有文件|*.*";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.DefaultExt = "doc";
            this.openFileDialog2.Filter = "所有word文档|*.doc; *.docx; *.dot; *.dotx|所有文件|*.*";
            // 
            // openFileDialog3
            // 
            this.openFileDialog3.DefaultExt = "doc";
            this.openFileDialog3.Filter = "所有word文档|*.doc; *.docx; *.dot; *.dotx|所有文件|*.*";
            // 
            // openFileDialog4
            // 
            this.openFileDialog4.FileName = "测试设计与执行阶段表格";
            // 
            // menu4
            // 
            this.menu4.Enabled = false;
            this.menu4.Items.Add(this.button3);
            this.menu4.Items.Add(this.menu5);
            this.menu4.Label = "测试设计与执行阶段";
            this.menu4.Name = "menu4";
            this.menu4.ShowImage = true;
            // 
            // button11
            // 
            this.button11.Label = "第一次就绪";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            // 
            // menu5
            // 
            this.menu5.Enabled = false;
            this.menu5.Items.Add(this.button11);
            this.menu5.Items.Add(this.button12);
            this.menu5.Label = "配置项动态测试表格";
            this.menu5.Name = "menu5";
            this.menu5.ShowImage = true;
            // 
            // button12
            // 
            this.button12.Enabled = false;
            this.button12.Label = "第二次就绪";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.OpenFileDialog openFileDialog3;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        private System.Windows.Forms.OpenFileDialog openFileDialog4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu4;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
    }

}
