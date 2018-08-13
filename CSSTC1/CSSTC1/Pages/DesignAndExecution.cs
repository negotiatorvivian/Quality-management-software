using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.FileProcessors.models;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;
using CSSTC1.InputProcessors;

namespace CSSTC1.Pages {
    public partial class DesignAndExecution : Form {
        private DesignAndExeProcessor processor = new DesignAndExeProcessor();
        List<ComboBox> comboBoxes = new List<ComboBox>();
        private bool jingtaifenxi = false;
        private bool daimashencha = false;
        private bool daimazoucha = false;
        private bool ceshijiucupingshen = true;
        #region  构造函数与界面操作
        public DesignAndExecution() {
            InitializeComponent();
            if(ContentFlags.wendangshencha > 0)
                this.checkedListBox1.SetItemChecked(0, true);
            if(ContentFlags.jingtaifenxi > 0){
                this.checkedListBox1.SetItemChecked(1, true);
                this.button3.Enabled = true;
                this.jingtaifenxi = true;
            }
            if(ContentFlags.daimazoucha > 0){
                this.checkedListBox1.SetItemChecked(2, true);
                this.button4.Enabled = true;
                this.daimazoucha = true;
            }
            if(ContentFlags.daimashencha > 0){
                this.checkedListBox1.SetItemChecked(3, true);
                this.button5.Enabled = true;
                this.daimashencha = true;
            }
            if(ContentFlags.peizhiceshi > 0){
                this.checkedListBox1.SetItemChecked(4, true);
                this.checkedListBox1.SetItemChecked(5, true);
            }
            if(ContentFlags.luojiceshi > 0)
                this.checkedListBox1.SetItemChecked(6, true);
            if(ContentFlags.xitongceshi > 0)
                this.checkedListBox1.SetItemChecked(7, true);
            if(ContentFlags.xitonghuiguiceshi > 0)
                this.checkedListBox1.SetItemChecked(8, true);
            
        }

        //取消键
        private void button2_Click(object sender, EventArgs e) {
            this.Hide();
        }
        
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e) {
            string times = this.comboBox3.Text;
            switch(times) {
                case "1次": {
                        this.label3.Enabled = false;
                        this.dateTimePicker5.Enabled = false;
                        foreach(ComboBox box in this.comboBoxes)
                            box.Enabled = false;
                        break;
                    }
                case "2次": {
                        this.label3.Enabled = true;
                        this.dateTimePicker5.Enabled = true;
                        foreach(ComboBox box in this.comboBoxes)
                            box.Enabled = true;
                        break;
                    }
                default:
                    break;

            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e) {
            //this.button1.Enabled = true;
        }
        

        //测试就绪表格
        public bool fill_csjx_table() {
            if(ContentFlags.peizhiceshi == 0) {
                this.tableLayoutPanel2.Visible = false;
                return true;
            }
            if(ContentFlags.pro_infos.Count == 0)
                return false;
            int row_index = 1;
            //this.tableLayoutPanel2.RowCount = ContentFlags.pro_infos.Count + 1;
            for(int i = 0; i < ContentFlags.pro_infos.Count - 1; i++) {
                this.tableLayoutPanel2.RowCount += 1;
                this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            }
            foreach(ProjectInfo file in ContentFlags.pro_infos) {
                System.Windows.Forms.Label label1 = new System.Windows.Forms.Label();
                label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                label1.Text = file.rj_mingcheng;
                label1.AutoSize = true;
                label1.Dock = System.Windows.Forms.DockStyle.Fill;


                System.Windows.Forms.ComboBox comboBox1 = new ComboBox();
                comboBox1.BackColor = System.Drawing.SystemColors.Control;
                comboBox1.Dock = System.Windows.Forms.DockStyle.Fill;
                comboBox1.FormattingEnabled = true;
                comboBox1.Items.AddRange(new object[] {
            "1",
            "2"});
                comboBox1.Size = new System.Drawing.Size(93, 21);
                comboBox1.Text = "1";
                comboBox1.Enabled = false;
                this.comboBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
                comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
                comboBoxes.Add(comboBox1);
                this.tableLayoutPanel2.Controls.Add(label1, 0, row_index);
                this.tableLayoutPanel2.Controls.Add(comboBox1, 1, row_index);
                row_index += 1;
            }
            return true;
        }
        
        //测试环境选择2时
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            Control control = (Control)sender;
            int index = this.tableLayoutPanel2.Controls.IndexOf(control) / this.tableLayoutPanel2.ColumnCount;
            string temp = this.tableLayoutPanel2.GetControlFromPosition(1, index).Text;
            if(temp.Equals("2")) {
                string software_name = this.tableLayoutPanel2.GetControlFromPosition(0, index).Text;
                PopUpWindow pop_up = new PopUpWindow(software_name);
                pop_up.Show();
                this.tableLayoutPanel2.GetControlFromPosition(1, index).Enabled = false;
            }
        }

        /************************展示静态分析 代码走查 代码审查表*********************************/
        private void button3_Click(object sender, EventArgs e) {
            PopUpStaticAnaChart chart1 = new PopUpStaticAnaChart("静态分析模块名", this.jingtaifenxi);
            chart1.Show();
            this.jingtaifenxi = false;
            if(!this.daimazoucha && !this.daimashencha && !this.jingtaifenxi && !this.ceshijiucupingshen)
                this.button1.Enabled = true;
        }

        private void button5_Click(object sender, EventArgs e) {
            PopUpCodeCheckChart chart1 = new PopUpCodeCheckChart("代码审查模块名", "代码审查范围",
                "dmsc_software_list", "dmsc_software_info", ContentFlags.dmsc_software_info, this.daimashencha);
            chart1.Show();
            this.daimashencha = false;
            if(!this.daimazoucha && !this.daimashencha && !this.jingtaifenxi && !this.ceshijiucupingshen)
                this.button1.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e) {
            PopUpCodeCheckChart chart1 = new PopUpCodeCheckChart("代码走查模块名", "代码走查范围",
                "dmzc_software_list", "dmzc_software_info", ContentFlags.dmzc_software_info, this.daimazoucha);
            chart1.Show();
            this.daimazoucha = false;
            if(!this.daimazoucha && !this.daimashencha && !this.jingtaifenxi && !this.ceshijiucupingshen)
                this.button1.Enabled = true;
        }
        /************************展示静态分析 代码走查 代码审查表*********************************/
   
        /****************************展示测试就绪界面*****************************/
        private void button6_Click(object sender, EventArgs e) {
            if(this.ceshijiucupingshen)
                this.fill_csjx_table();
            this.panel1.Show();
            this.ceshijiucupingshen = false;
            if(!this.daimazoucha && !this.daimashencha && !this.jingtaifenxi && !this.ceshijiucupingshen)
                this.button1.Enabled = true;
        }

        public bool fill_table() {
            if(ContentFlags.static_files.Count == 0)
                return true;
            //return false;
            if(TimeStamp.sldtcs_time != null) {
                DateTime date1 = DateHelper.cal_date(TimeStamp.sldtcs_time, 10);
                DateTime date2 = DateHelper.cal_date(TimeStamp.sldtcs_time, 5);
                this.dateTimePicker1.Value = date1;
                this.dateTimePicker4.Value = date2;
                this.dateTimePicker5.Value = date2;
            }
            return true;
        }
        /************************展示测试就绪界面*********************************/
        #endregion

        //确认键提交信息
        private void button1_Click(object sender, EventArgs e) {
            this.button1.Enabled = false;
            TimeStamp.cssmps_time = this.dateTimePicker1.Value.ToShortDateString();
            TimeStamp.cssmps_format_time = this.dateTimePicker1.Value.ToLongDateString();
            this.Hide();
        }

        private void button7_Click(object sender, EventArgs e) {
            string csjx_questions = this.comboBox2.Text;
            TimeStamp.csjxps_time.Add(this.dateTimePicker4.Value.ToShortDateString());
            TimeStamp.csjxps_format_time.Add(this.dateTimePicker4.Value.ToLongDateString());
            if(this.dateTimePicker5.Enabled){
                TimeStamp.csjxps_time.Add(this.dateTimePicker5.Value.ToShortDateString());
                TimeStamp.csjxps_format_time.Add(this.dateTimePicker5.Value.ToLongDateString());
            }
            if(csjx_questions.Equals("无"))
                ContentFlags.pianli_3 = 0;
            bool flag = false;
            foreach(ComboBox box in comboBoxes){
                if(box.Text.Equals("2")){
                    this.processor.fill_time_line(); 
                    this.panel1.Hide();
                    flag = true;
                    break;
                }
            }
            if(!flag)
                MessageBox.Show("请填写第二次就绪软件的环境信息！");
        }

        
    }
}
