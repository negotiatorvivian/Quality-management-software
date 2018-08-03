using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.CommonUtils;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.readers;
using CSSTC1.InputProcessors;

namespace CSSTC1.Pages {
    public partial class DemandAnalysis : Form {
        private FileReader2 reader1 = new FileReader2();
        private FileReader3 reader2 = new FileReader3();
        private DemandAnaProcessor processor = new DemandAnaProcessor();

        #region  页面组件行为控制
        public DemandAnalysis() {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e) {
            Globals.ThisDocument.demand_analysis_form.Hide();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            this.panel1.Show();
        }

        private void button3_Click(object sender, EventArgs e) {
            this.panel1.Hide();
        }

        private void button4_Click(object sender, EventArgs e) {
            if(this.textBox1.Text.Length != 0) {
                object a = this.textBox1.Text;
                this.checkedListBox1.Items.Add(a, true);
            }
            this.textBox1.Text = "";
            this.panel1.Hide();
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e) {
            string times = this.comboBox2.Text;
            switch(times) {
                case "有": {
                        this.label5.Enabled = true;
                        this.dateTimePicker5.Enabled = true;
                        break;
                    }
                case "无": {
                        this.label5.Enabled = false;
                        this.dateTimePicker5.Enabled = false;
                        break;
                    }
                default:
                    break;
            }
        }
        #endregion

        //确认提交
        private void button1_Click(object sender, EventArgs e) {
            this.button1.Enabled = false;
            System.Windows.Forms.CheckedListBox.CheckedItemCollection name_list = 
                          this.checkedListBox1.CheckedItems;
            List<string> names = new List<string>();
            
            for(int i = 0; i < name_list.Count; i++){
                string temp = name_list[i].ToString();
                names.Add(temp);
            }
            string ceshixuqiu_time = this.xuqiufenxi_shijian.Value.ToLongDateString();
            TimeStamp.ceshins_time = this.xuqiufenxi_shijian.Text;
            if(ContentFlags.ceshidagang)
                reader1.read_psz_members(names);
            else
                reader2.read_psz_members(names);
            string pianli2_time = "";
            if(this.dateTimePicker5.Enabled){
                TimeStamp.pianli2_time = this.dateTimePicker5.Text;
                pianli2_time = this.dateTimePicker5.Value.ToLongDateString();
                processor.fill_time_info(ceshixuqiu_time, pianli2_time);
            }
            else{
                ContentFlags.pianli_2 = 0;
                processor.fill_time_info(TimeStamp.ceshins_time, pianli2_time);

            }
            Globals.ThisDocument.demand_analysis_form.Hide();
        }

        //设置默认值
        public void change_status() {
            if(ContentFlags.ceshidagang) {
                this.checkBox1.Checked = true;
                this.checkBox2.Enabled = false;
            }
            else {
                this.checkBox2.Checked = true;
                this.checkBox1.Enabled = false;
            }
            this.checkedListBox1.SetItemChecked(0, true);
            for(int i = 1; i < 5; i++) {
                Random r = new Random();
                if(r.NextDouble() > 0.5 && i != 4) {
                    this.checkedListBox1.SetItemChecked(2 * i, true);
                }
                else
                    this.checkedListBox1.SetItemChecked(2 * i - 1, true);
            }
            if(TimeStamp.ceshisc_time != null) {
                DateTime neishenshijian = DateHelper.cal_date(TimeStamp.ceshisc_time, 14);
                this.xuqiufenxi_shijian.Value = neishenshijian;
            }

        }
    }
}
