using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.InputProcessors;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;

namespace CSSTC1.Pages {
    public partial class ProjectEstabInfo : Form {
        public ProjectEstabInfoProcessor processor = new ProjectEstabInfoProcessor();
        public ProjectEstabInfo() {
            InitializeComponent();
        }
        #region  用户界面组件控制
        private void button2_Click(object sender, EventArgs e) {
            Globals.ThisDocument.project_estab_info.Hide();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e) {
            string times = this.comboBox1.Text;
            switch(times){
                case "2次":{
                    this.label5.Enabled = true;
                    this.dateTimePicker2.Enabled = true;
                    this.label6.Enabled = false;
                    this.dateTimePicker3.Enabled = false;
                    this.label7.Enabled = false;
                    this.dateTimePicker4.Enabled = false;
                    break;
                }
                case "3次": {
                        this.label5.Enabled = true;
                        this.dateTimePicker2.Enabled = true;
                        this.label6.Enabled = true;
                        this.dateTimePicker3.Enabled = true;
                        this.label7.Enabled = false;
                        this.dateTimePicker4.Enabled = false;
                        break;
                    }
                case "4次": {
                        this.label5.Enabled = true;
                        this.dateTimePicker2.Enabled = true;
                        this.label6.Enabled = true;
                        this.dateTimePicker3.Enabled = true;
                        this.label7.Enabled = true;
                        this.dateTimePicker4.Enabled = true;
                        break;
                    }
                default:{
                        this.label5.Enabled = false;
                        this.dateTimePicker2.Enabled = false;
                        this.label6.Enabled = false;
                        this.dateTimePicker3.Enabled = false;
                        this.label7.Enabled = false;
                        this.dateTimePicker4.Enabled = false;
                        break;
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) {
            string times = this.comboBox2.Text;
            switch(times) {
                case "有": {
                        this.label4.Enabled = true;
                        this.dateTimePicker5.Enabled = true;
                        break;
                    }
                case "无": {
                        this.label4.Enabled = false;
                        this.dateTimePicker5.Enabled = false;
                        break;
                    }
                default:
                    break;
            }
        }
        #endregion

        //提交信息
        private void button1_Click(object sender, EventArgs e) {
            this.button1.Enabled = false;
            string lx_time = this.numericUpDown1.Value.ToString();
            //decimal lx_time = this.numericUpDown1.Value;
            string lq_times = this.comboBox1.Text;
            List<DateTime> lq_time = new List<DateTime>();
            if(lq_times[0].Equals('1')){
                ContentFlags.lingqucishu = 1;
                lq_time.Add(this.dateTimePicker1.Value);
                TimeStamp.lingqushijian.Add(this.dateTimePicker1.Text);
            }
            else if(lq_times[0].Equals('2')) {
                ContentFlags.lingqucishu = 2;
                lq_time.Add(this.dateTimePicker1.Value);
                lq_time.Add(this.dateTimePicker2.Value);
                TimeStamp.lingqushijian.Add(this.dateTimePicker1.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker2.Text);
            }
            else if(lq_times[0].Equals('3')) {
                ContentFlags.lingqucishu = 3;
                lq_time.Add(this.dateTimePicker1.Value);
                lq_time.Add(this.dateTimePicker2.Value);
                lq_time.Add(this.dateTimePicker3.Value);
                TimeStamp.lingqushijian.Add(this.dateTimePicker1.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker2.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker3.Text);
            }
            else{
                ContentFlags.lingqucishu = lq_times[0];
                lq_time.Add(this.dateTimePicker1.Value);
                lq_time.Add(this.dateTimePicker2.Value);
                lq_time.Add(this.dateTimePicker3.Value);
                lq_time.Add(this.dateTimePicker4.Value);
                TimeStamp.lingqushijian.Add(this.dateTimePicker1.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker2.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker3.Text);
                TimeStamp.lingqushijian.Add(this.dateTimePicker4.Text);
            }
            DateTime pl_time = DateTime.MaxValue;
            if(this.dateTimePicker5.Enabled == true){
                pl_time = this.dateTimePicker5.Value;
                //pl_time = DateHelper.cal_date(temp, 0);
            }
            processor.fill_estab_info(lx_time, lq_time, pl_time);
            Globals.ThisDocument.project_estab_info.Hide();
        }

    }
}
