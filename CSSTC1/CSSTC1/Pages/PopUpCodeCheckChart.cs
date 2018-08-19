using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;
using System.Text.RegularExpressions;
using System.Reflection;

namespace CSSTC1.Pages {
    public partial class PopUpCodeCheckChart : Form {
        public PopUpCodeCheckChart(string name, string label, string var1, string var2,
            List<StaticAnalysisFile> files, bool enabled) {
            InitializeComponent();
            this.Text = name;
            this.label3.Text = label;
            this.files = files;
            this.var1 = var1;
            this.var2 = var2;
            this.fill_table();
        }
        private string var1;
        private string var2;
        List<LinkLabel> labels = new List<LinkLabel>();
        List<StaticAnalysisFile> files;

        #region 在界面窗体动态创建更新静态项表格
        public bool fill_table() {
            if(files.Count == 0)
                return false;
            List<System.Windows.Forms.TextBox> boxes = new List<TextBox>();
            List<string> huigui_banben = this.update_version();
            int row_index = 1;
            //this.tableLayoutPanel1.RowCount = files.Count + 1;
            for(int i = 0; i < files.Count - 1; i++) {
                this.tableLayoutPanel1.RowCount += 1;
                this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            }
            foreach(StaticAnalysisFile file in files) {
                System.Windows.Forms.Label label1 = new System.Windows.Forms.Label();
                label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                label1.Text = file.rj_mingcheng;
                label1.AutoSize = true;
                label1.Dock = System.Windows.Forms.DockStyle.Fill;

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

                System.Windows.Forms.TextBox textBox4 = new System.Windows.Forms.TextBox();
                textBox4.BackColor = System.Drawing.SystemColors.Control;
                textBox4.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox4.MaximumSize = new System.Drawing.Size(100, 30);
                textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox4.Margin = new System.Windows.Forms.Padding(0);
                textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox4.AutoSize = true;
                textBox4.Multiline = true;
                textBox4.Text = file.xt_banben.Split('/')[0];

                System.Windows.Forms.TextBox textBox6 = new System.Windows.Forms.TextBox();
                textBox6.BackColor = System.Drawing.SystemColors.Control;
                textBox6.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox6.MaximumSize = new System.Drawing.Size(100, 30);
                textBox6.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox6.Margin = new System.Windows.Forms.Padding(0);
                textBox6.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox6.AutoSize = true;
                textBox6.Multiline = true;
                textBox6.Text = huigui_banben[row_index - 1];

                System.Windows.Forms.TextBox textBox5 = new System.Windows.Forms.TextBox();
                textBox5.BackColor = System.Drawing.SystemColors.Control;
                textBox5.Dock = System.Windows.Forms.DockStyle.Fill;
                textBox5.MaximumSize = new System.Drawing.Size(157, 30);
                textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
                textBox5.Margin = new System.Windows.Forms.Padding(0);
                textBox5.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                textBox5.AutoSize = true;
                textBox5.Multiline = true;
                textBox5.Text = file.yz_danwei;

                LinkLabel linkLabel1 = this.add_delete_label();
                this.tableLayoutPanel1.Controls.Add(label1, 0, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox2, 1, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox4, 2, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox6, 3, row_index);
                this.tableLayoutPanel1.Controls.Add(textBox5, 4, row_index);
                this.tableLayoutPanel1.Controls.Add(linkLabel1, 5, row_index);
                row_index += 1;
            }
            return true;
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e) {
            //int index = this.labels.IndexOf((LinkLabel)sender);
            Control control = (Control)sender;
            int index = this.tableLayoutPanel1.Controls.IndexOf(control) / this.tableLayoutPanel1.ColumnCount;
            for(int i = 0; i < this.tableLayoutPanel1.ColumnCount; i++) {
                this.tableLayoutPanel1.Controls.Remove(this.tableLayoutPanel1.Controls[(index) * 6]);
            }
            this.tableLayoutPanel1.Refresh();

        }

        private LinkLabel add_delete_label() {
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

        private List<string> update_version() {
            List<string> huigui_banben = new List<string>();
            foreach(StaticAnalysisFile file in this.files) {
                string version = file.xt_banben;
                Regex reg = new Regex("[v|V][0-9][.][0-9]+");
                MatchCollection matches = reg.Matches(version, 0);
                string temp = matches[0].Value;
                string temp1 = temp.Substring(1, temp.Length - 1);
                double t = Double.Parse(temp1);
                t = t + 0.1;
                version = "V" + t.ToString();
                file.hg_banben = version;
                huigui_banben.Add(version);
            }
            return huigui_banben;
        }
        #endregion

        private void button1_Click(object sender, EventArgs e) {
            this.read_current_list();
            this.Hide();
        }

        //读取代码走查或代码审查模块名页面上更改后的信息
        public bool read_current_list() {
            TableLayoutControlCollection controls = this.tableLayoutPanel1.Controls;
            List<string> software_list = new List<string>();
            List<StaticAnalysisFile> files = new List<StaticAnalysisFile>();
            for(int i = 6; i < controls.Count; i += 6) {
                string name = controls[i].Text;
                software_list.Add(name);
                string range = controls[i + 1].Text;
                string old_version = controls[i + 2].Text;
                string new_version = controls[i + 3].Text;
                string provider = controls[i + 4].Text;
                
                StaticAnalysisFile file = new StaticAnalysisFile(name, range, old_version, new_version, "",
                    provider);
                files.Add(file);
            }
            BindingFlags flag = BindingFlags.Static | BindingFlags.Public;
            FieldInfo f_key = typeof(ContentFlags).GetField(this.var1, flag);
            List<String> temp = new List<string>();
            f_key.SetValue(temp, software_list);
            List<StaticAnalysisFile> new_files = new List<StaticAnalysisFile>();
            FieldInfo f_key1 = typeof(ContentFlags).GetField(this.var2, flag);
            f_key1.SetValue(new_files, files);
            return true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            if(this.Text.Equals("代码走查模块名")) {
                ContentFlags.dmzc_same = this.checkBox1.Checked;
            }
            else
                ContentFlags.dmsc_same = this.checkBox1.Checked;
        }
    }
}
