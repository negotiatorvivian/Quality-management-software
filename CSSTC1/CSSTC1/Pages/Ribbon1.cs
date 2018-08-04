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
        }

        //填写项目立项阶段信息
        private void button5_Click(object sender, RibbonControlEventArgs e) {
            this.project_estab_info.show_estab_info();
        }

        //读取测试大纲或需求说明文件
        private void button10_Click(object sender, RibbonControlEventArgs e) {
            if(!ContentFlags.ceshidagang) {
                if(openFileDialog3.ShowDialog() == DialogResult.OK) {
                    string read_in_file = openFileDialog3.FileName;
                    file_reader3.read_charts(read_in_file);
                }
            }
            else
                MessageBox.Show("立项阶段没有选择需求规格说明！");
        }

        //填写测试大纲或需求说明信息
        private void button9_Click(object sender, RibbonControlEventArgs e) {
            if(ContentFlags.ceshidagang) {
                if(openFileDialog2.ShowDialog() == DialogResult.OK) {
                    string read_in_file = openFileDialog2.FileName;
                    file_reader2.read_charts(read_in_file);
                }
            }
            else
                MessageBox.Show("立项阶段没有选择测试大纲！");
        }

        //根据已填写信息更新测试大纲或需求说明阶段界面的默认值
        private void button6_Click(object sender, RibbonControlEventArgs e) {
            demand_analysis_form.change_status();
            this.demand_analysis_form.Show();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e) {
            OperationHelper.update_file();
            MessageBox.Show("文档更新完成!");
        }

        private void button7_Click(object sender, RibbonControlEventArgs e) {
            bool res = Globals.ThisDocument.design_and_exe.fill_table();
            if(res)
                Globals.ThisDocument.design_and_exe.Show();
        }

        
    }
}
