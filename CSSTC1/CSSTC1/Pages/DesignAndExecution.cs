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

        public DesignAndExecution() {
            InitializeComponent();
            this.checkedListBox1.SetItemChecked(0, true);
            this.checkedListBox1.SetItemChecked(1, true);
            this.checkedListBox1.SetItemChecked(4, true);
            this.checkedListBox1.SetItemChecked(5, true);
            this.checkedListBox1.SetItemChecked(6, true);
            this.checkedListBox1.SetItemChecked(7, true);
            this.checkedListBox1.SetItemChecked(8, true);
            
        }

        public bool fill_table() {
            if(ContentFlags.static_files.Count == 0)
                return false;
            if(TimeStamp.sldtcs_time != null) {
                DateTime date1 = DateHelper.cal_date(TimeStamp.sldtcs_time, 10);
                DateTime date2 = DateHelper.cal_date(TimeStamp.sldtcs_time, 5);
                this.dateTimePicker1.Value = date1;
                this.dateTimePicker2.Value = date2;
            }
            return true;
        }

        ////读取页面上更改后的信息
        //public bool read_current_list(){
        //    TableLayoutControlCollection controls = this.tableLayoutPanel1.Controls;
        //    List<string> software_list = new List<string>();
        //    for(int i = 7; i < controls.Count; i +=7){
        //        string name = controls[i].Text;
        //        software_list.Add(name);
        //    }
        //    if(software_list.Count == 0){
        //        return false;
        //    }
        //    ContentFlags.software_list = software_list;
        //    //读取修改过的信息
        //    return true;
        //}

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
        //确认键提交信息
        private void button1_Click(object sender, EventArgs e) {
            this.button1.Enabled = false;
            string csjxps_format_time = "";
            if(this.dateTimePicker3.Enabled)
                csjxps_format_time = this.dateTimePicker3.Value.ToLongDateString();
            //bool res1 = this.read_current_list();
            //if(!res1)
            //    MessageBox.Show("未选择静态测试软件项目");

            //bool res = this.processor.fill_time_line(csjxps_format_time);
            //if(!res)
            //    MessageBox.Show("填写测试设计与执行阶段时间线出错");
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e) {
            Globals.ThisDocument.design_and_exe.Hide();
        }

        private void button3_Click(object sender, EventArgs e) {
            PopUpStaticAnaChart chart1 = new PopUpStaticAnaChart("静态分析模块名");
            chart1.Show();
        }
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e) {
            TimeStamp.cssmps_time = this.dateTimePicker1.Value.ToShortDateString();
            TimeStamp.cssmps_format_time = this.dateTimePicker1.Value.ToLongDateString();
            //foreach(ComboBox box in comboBoxes) {
            //    box.Enabled = true;
            //}
            this.button1.Enabled = true;
            DateTime temp1 = DateHelper.cal_date(TimeStamp.csjxps_time, 1);
            DateTime temp2 = DateHelper.cal_date(TimeStamp.csjxps_time, 1);
        }

        private void button2_Click_1(object sender, EventArgs e) {
            this.Hide();
        }

        private void button5_Click(object sender, EventArgs e) {
            PopUpCodeCheckChart chart1 = new PopUpCodeCheckChart("代码审查模块名", "代码审查范围", 
                "dmsc_software_list", "dmsc_software_info", ContentFlags.dmsc_software_info);
            chart1.Show();
        }

        private void button4_Click(object sender, EventArgs e) {
            PopUpCodeCheckChart chart1 = new PopUpCodeCheckChart("代码走查模块名", "代码走查范围", 
                "dmzc_software_list", "dmzc_software_info", ContentFlags.dmzc_software_info);
            chart1.Show();
        }
    }
}
