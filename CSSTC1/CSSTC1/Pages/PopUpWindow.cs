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
        
        public PopUpWindow(string software_name) {
            InitializeComponent();
            this.software_name = software_name;
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

    }
}
