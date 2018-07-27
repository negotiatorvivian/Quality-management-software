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

namespace CSSTC1.Pages {
    public partial class ProjectEstabInfo : Form {
        public ProjectEstabInfoProcessor processor = new ProjectEstabInfoProcessor();
        public ProjectEstabInfo() {
            InitializeComponent();
        }

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

        private void button1_Click(object sender, EventArgs e) {
            string lx_time = this.numericUpDown1.Value.ToString();
            //decimal lx_time = this.numericUpDown1.Value;
            string lq_times = this.comboBox1.Text;
            List<string> lq_time = new List<string>();
            if(lq_times[0].Equals('1')){
                lq_time.Add(this.dateTimePicker1.Text);
            }
            else if(lq_times[0].Equals('2')) {
                ContentFlags.lingqucishu = 2;
                lq_time.Add(this.dateTimePicker1.Text);
                lq_time.Add(this.dateTimePicker2.Text);
            }
            else if(lq_times[0].Equals('3')) {
                ContentFlags.lingqucishu = 3;
                lq_time.Add(this.dateTimePicker1.Text);
                lq_time.Add(this.dateTimePicker2.Text);
                lq_time.Add(this.dateTimePicker3.Text);
            }
            else{
                ContentFlags.lingqucishu = lq_times[0];
                lq_time.Add(this.dateTimePicker1.Text);
                lq_time.Add(this.dateTimePicker2.Text);
                lq_time.Add(this.dateTimePicker3.Text);
                lq_time.Add(this.dateTimePicker4.Text);
            }
            string pl_time = "";
            if(this.dateTimePicker5.Enabled == true){
                pl_time = this.dateTimePicker5.Text;
            }
            processor.fill_estab_info(lx_time, lq_time, pl_time);
            Globals.ThisDocument.project_estab_info.Hide();
        }

        

     
    }
}
