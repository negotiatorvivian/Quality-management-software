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

namespace CSSTC1.Pages {
    public partial class PopUpWindow : Form {
        public string software_name = "";
        private List<string> software_names = ContentFlags.pro_infos.Select(r => r.rj_mingcheng).ToList();
        private List<string> software_providers = ContentFlags.pro_infos.Select(r => r.yz_danwei).ToList();
        
        public PopUpWindow(string software_name, DateTime date1, DateTime date2) {
            InitializeComponent();
            this.software_name = software_name;
            this.dateTimePicker1.Value = date1;
            this.dateTimePicker3.Value = date2;
            this.initialize_location();
        }

        private void button1_Click(object sender, EventArgs e) {
            string time1 = this.dateTimePicker1.Value.ToLongDateString();
            string time2 = this.dateTimePicker3.Value.ToLongDateString();
            string local1 = this.textBox1.Text;
            string local2 = this.textBox2.Text;
            TestEnvironment test_env = new TestEnvironment(this.software_name, time1, time2, local1, local2);
            ContentFlags.test_envs.Add(test_env);
            this.button1.Enabled = false;
            this.Hide();
        }

        private void initialize_location(){
            int index = this.software_names.IndexOf(this.software_name);
            if(index >= 0){
                this.textBox2.Text = this.software_providers[index];
            }
        }
    }
}
