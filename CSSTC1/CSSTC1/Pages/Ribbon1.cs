using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Aspose.Words;

using CSSTC1.FileProcessors;
using CSSTC1.FileProcessors.readers;
using CSSTC1.InputProcessors;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;

namespace CSSTC1.Pages {
    public partial class Ribbon1 {
        public FileReader1 file_reader = new FileReader1();
        public FileReader2 file_reader2 = new FileReader2();
        public FileReader3 file_reader3 = new FileReader3();
        public FileReader4 file_reader4 = new FileReader4();
        public FileReader5 file_reader5;
        public ProjectEstabInfoProcessor project_estab_info = new ProjectEstabInfoProcessor();
        public DemandAnalysis demand_analysis_form;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {
            this.demand_analysis_form = Globals.ThisDocument.demand_analysis_form;
            
        }

        //读取项目立项阶段文件
        private void button1_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog1.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog1.FileName;
                file_reader.read_charts(read_in_file);
            }
            this.button1.Enabled = false;
            //if(!this.button6.Enabled)
            this.menu1.Enabled = true;

        }

        //填写项目立项阶段信息
        private void button5_Click(object sender, RibbonControlEventArgs e) {
            this.project_estab_info.show_estab_info();
            this.button6.Enabled = true;
            //this.button1.Enabled = true;
            this.button5.Enabled = false;
        }

        //读取测试大纲文件
        private void button9_Click(object sender, RibbonControlEventArgs e) {
            if(ContentFlags.ceshidagang) {
                if(openFileDialog2.ShowDialog() == DialogResult.OK) {
                    string read_in_file = openFileDialog2.FileName;
                    file_reader2.read_charts(read_in_file);
                }
                this.menu1.Enabled = false;
                this.button7.Enabled = true;
            }
            else
                MessageBox.Show("立项阶段没有选择测试大纲！");
        }

        //读取需求说明文件
        private void button10_Click(object sender, RibbonControlEventArgs e) {
            if(!ContentFlags.ceshidagang) {
                if(openFileDialog3.ShowDialog() == DialogResult.OK) {
                    string read_in_file = openFileDialog3.FileName;
                    file_reader3.read_charts(read_in_file);
                }
                this.menu1.Enabled = false;
                this.button7.Enabled = true;
            }
            else
                MessageBox.Show("立项阶段没有选择需求规格说明！");
        }


        //读取测试设计与执行阶段文件
        private void button3_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog4.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog4.FileName;
                file_reader4.read_charts(read_in_file);
            } 
            this.button3.Enabled = false;
            this.menu5.Enabled = true;
        }

        //根据已填写信息更新测试大纲或需求说明阶段界面的默认值
        private void button6_Click(object sender, RibbonControlEventArgs e) {
            demand_analysis_form.change_status();
            this.demand_analysis_form.Show();
            //if(!this.button1.Enabled)
            //    this.menu1.Enabled = true;
            this.button1.Enabled = true;
            this.button6.Enabled = false;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e) {
            OperationHelper.update_file();
            MessageBox.Show("文档更新完成!");
        }

        private void button7_Click(object sender, RibbonControlEventArgs e) {
            Globals.ThisDocument.design_and_exe = new DesignAndExecution();
            bool res = Globals.ThisDocument.design_and_exe.fill_table();
            if(res)
                Globals.ThisDocument.design_and_exe.Show();
            this.button7.Enabled = false;
            this.menu4.Enabled = true;
        }

        private void button11_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog5.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog5.FileName;
                file_reader5 = new FileReader5(false);
                file_reader5.read_charts(read_in_file);
            }
            this.button11.Enabled = false;
            this.button12.Enabled = true;
        }

        private void button12_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog6.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog6.FileName;
                file_reader5 = new FileReader5(true);
                file_reader5.read_charts(read_in_file);
            }
            this.button12.Enabled = false;
            this.button8.Enabled = true;
        }

        

        
    }
}
