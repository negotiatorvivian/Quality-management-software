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
        List<ComboBox> comboBoxes = new List<ComboBox>();
        List<LinkLabel> labels = new List<LinkLabel>();
        private DesignAndExeProcessor processor = new DesignAndExeProcessor();

        public DesignAndExecution() {
            InitializeComponent();
            this.checkedListBox1.SetItemChecked(0, true);
            this.checkedListBox1.SetItemChecked(1, true);
            this.checkedListBox1.SetItemChecked(4, true);
            this.checkedListBox1.SetItemChecked(5, true);
            this.checkedListBox1.SetItemChecked(6, true);
            this.checkedListBox1.SetItemChecked(7, true);
            this.checkedListBox1.SetItemChecked(8, true);
            if(TimeStamp.sldtcs_time != null){
                DateTime date1 = DateHelper.cal_date(TimeStamp.sldtcs_time, 10);
                DateTime date2 = DateHelper.cal_date(TimeStamp.sldtcs_time, 5);
                this.dateTimePicker1.Value = date1;
                this.dateTimePicker2.Value = date2;

            }
        }

        #region 在界面窗体动态创建更新静态项表格
        public bool fill_table(){
            if(ContentFlags.static_files.Count == 0)
                return false;
            List<System.Windows.Forms.TextBox> boxes = new List<TextBox>();
            int row_index = 1;
            this.tableLayoutPanel1.RowCount = ContentFlags.static_files.Count + 1;
            foreach(StaticAnalysisFile file in ContentFlags.static_files) {
                System.Windows.Forms.Label label1 = new System.Windows.Forms.Label();
                label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                label1.Text = file.rj_mingcheng;
                label1.AutoSize = true;
                label1.Dock = System.Windows.Forms.DockStyle.Fill;

                //System.Windows.Forms.Label label2 = new System.Windows.Forms.Label();
                //label2.AutoSize = true;
                //label2.Dock = System.Windows.Forms.DockStyle.Fill;
                //label2.Size = new System.Drawing.Size(94, 30);
                //label2.Text = file.jtfx_fanwei;
                //label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

                System.Windows.Forms.TextBox textBox2 = new System.Windows.Forms.TextBox();
                textBox2.BackColor = System.Drawing.SystemColors.Control;
                textBox2.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox2.MaximumSize = new System.Drawing.Size(94, 30);
                textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox2.Margin = new System.Windows.Forms.Padding(0);
                textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox2.AutoSize = true;
                textBox2.Multiline = true;
                textBox2.Text = file.jtfx_fanwei;

                System.Windows.Forms.TextBox textBox3 = new System.Windows.Forms.TextBox();
                textBox3.BackColor = System.Drawing.SystemColors.Control;
                textBox3.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox3.MaximumSize = new System.Drawing.Size(100, 0);
                textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox3.Margin = new System.Windows.Forms.Padding(0);
                textBox3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox3.AutoSize = true;
                textBox3.Multiline = true;
                textBox3.Text = file.bc_yuyan;

                //System.Windows.Forms.Label label4 = new System.Windows.Forms.Label();
                //label4.AutoSize = true;
                //label4.Dock = System.Windows.Forms.DockStyle.Fill;
                //label4.Size = new System.Drawing.Size(100, 30);
                //label4.Text = file.xt_banben;
                //label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                System.Windows.Forms.TextBox textBox4 = new System.Windows.Forms.TextBox();
                textBox4.BackColor = System.Drawing.SystemColors.Control;
                textBox4.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox4.MaximumSize = new System.Drawing.Size(100, 30);
                textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox4.Margin = new System.Windows.Forms.Padding(0);
                textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox4.AutoSize = true;
                textBox4.Multiline = true;
                textBox4.Text = file.xt_banben;

                //System.Windows.Forms.Label label5 = new System.Windows.Forms.Label();
                //label5.AutoSize = true;
                //label5.Dock = System.Windows.Forms.DockStyle.Fill;
                //label5.Size = new System.Drawing.Size(176, 30);
                //label5.Text = file.yz_danwei;
                //label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                System.Windows.Forms.TextBox textBox5 = new System.Windows.Forms.TextBox();
                textBox5.BackColor = System.Drawing.SystemColors.Control;
                textBox5.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox5.MaximumSize = new System.Drawing.Size(176, 30);
                textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox5.Margin = new System.Windows.Forms.Padding(0);
                textBox5.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox5.AutoSize = true;
                textBox5.Multiline = true;
                textBox5.Text = file.yz_danwei;
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
                comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
                comboBoxes.Add(comboBox1);
                LinkLabel linkLabel1 = this.add_delete_label();
                this.tableLayoutPanel1.Controls.Add(label1, 0, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox2, 1, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox3, 2, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox4, 3, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox5, 4, row_index);
                this.tableLayoutPanel1.Controls.Add(comboBox1, 5, row_index);
                this.tableLayoutPanel1.Controls.Add(linkLabel1, 6, row_index);
                row_index += 1;
            }
            return true;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            Control control = (Control)sender;
            int index = this.tableLayoutPanel1.Controls.IndexOf(control) / this.tableLayoutPanel1.ColumnCount;
            string temp = this.tableLayoutPanel1.GetControlFromPosition(5, index).Text;
            if(temp.Equals("2")){
                string software_name = this.tableLayoutPanel1.GetControlFromPosition(0, index).Text;
                PopUpWindow pop_up = new PopUpWindow(software_name);
                pop_up.Show();
                this.tableLayoutPanel1.GetControlFromPosition(5, index).Enabled = false;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e) {
            TimeStamp.cssmps_time = this.dateTimePicker1.Value.ToShortDateString();
            TimeStamp.cssmps_format_time = this.dateTimePicker1.Value.ToLongDateString();
            foreach(ComboBox box in comboBoxes){
                box.Enabled = true;
            }
            this.button1.Enabled = true;
            DateTime temp1 = DateHelper.cal_date(TimeStamp.csjxps_time, 1);
            DateTime temp2 = DateHelper.cal_date(TimeStamp.csjxps_time, 1);
        }

        private void button2_Click_1(object sender, EventArgs e) {
            this.Hide();
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            //int index = this.labels.IndexOf((LinkLabel)sender);
            Control control = (Control)sender;
            int index = this.tableLayoutPanel1.Controls.IndexOf(control) / this.tableLayoutPanel1.ColumnCount;
            for(int i = 0 ; i < this.tableLayoutPanel1.ColumnCount; i ++){
                this.tableLayoutPanel1.Controls.Remove(this.tableLayoutPanel1.Controls[(index) * 7]);
            }
            this.tableLayoutPanel1.Refresh();

        }

        private LinkLabel add_delete_label(){
            // 
            // linkLabel1
            // 
            LinkLabel linkLabel1 = new LinkLabel();
            linkLabel1.Anchor = System.Windows.Forms.AnchorStyles.None;
            linkLabel1.AutoSize = true;
            linkLabel1.Size = new System.Drawing.Size(35, 14);
            linkLabel1.TabIndex = 7;
            linkLabel1.TabStop = true;
            linkLabel1.Text = "删除";
            linkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(linkLabel1_LinkClicked);
            this.labels.Add(linkLabel1);
            return linkLabel1;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) {
            string times = this.comboBox2.Text;
            switch(times) {
                case "有": {
                        this.label8.Enabled = true;
                        this.dateTimePicker3.Enabled = true;
                        break;
                    }
                case "无": {
                        this.label8.Enabled = false;
                        this.dateTimePicker3.Enabled = false;
                        break;
                    }
                default:
                    break;
            }
        }
        #endregion

        //读取页面上更改后的信息
        public bool read_current_list(){
            TableLayoutControlCollection controls = this.tableLayoutPanel1.Controls;
            List<string> software_list = new List<string>();
            for(int i = 7; i < controls.Count; i +=7){
                string name = controls[i].Text;
                software_list.Add(name);
            }
            if(software_list.Count == 0){
                return false;
            }
            ContentFlags.software_list = software_list;
            //读取修改过的信息


            return true;
        }


        //确认键提交信息
        private void button1_Click(object sender, EventArgs e) {
            this.button1.Enabled = false;
            string csjxps_format_time = "";
            if(this.dateTimePicker3.Enabled)
                csjxps_format_time = this.dateTimePicker3.Value.ToLongDateString();
            bool res1 = this.read_current_list();
            if(!res1)
                MessageBox.Show("未选择静态测试软件项目");

            bool res = this.processor.fill_time_line(csjxps_format_time);
            if(!res)
                MessageBox.Show("填写测试设计与执行阶段时间线出错");
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e) {
            Globals.ThisDocument.design_and_exe.Hide();
        }
    }
}
