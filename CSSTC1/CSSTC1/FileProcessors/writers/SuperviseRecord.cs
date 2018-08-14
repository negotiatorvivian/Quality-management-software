using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using System.Text.RegularExpressions;
using Aspose.Words;
using CSSTC1.CommonUtils;
using System.Windows;
using CSSTC1.FileProcessors.writers.BuildEnvironment;

namespace CSSTC1.FileProcessors.writers {
    class SuperviseRecord {
        private int current_row_index = 1;
        List<string> contents = new List<string>();
        List<String> bookmarks = new List<string>();
        List<string> keys = new List<string>();

        public SuperviseRecord(){
            this.keys = ContentFlags.sys_time_dict.Select(r => r.Key).ToList();
            //this.values = ContentFlags.sys_time_dict.Select(r => r.Value).ToList();
            this.write_time_line();
        }

        public void write_time_line(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            doc_builder.MoveToBookmark("文档审查阶段时间");
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, 0, true);
            Dictionary<string, int> problem_sum = this.count_problem();
            int current_table_index = 1;
            if(ContentFlags.wendangshencha > 0){
                this.wdsc_stage_sum(problem_sum["文档审查"]);
                current_table_index += 1;
            }
            else{
                this.remove_rows(table, this.current_row_index, 3);
                OperationHelper.delete_table(doc, doc_builder, "文档审查SQA", current_table_index);
            }
            if(ContentFlags.jingtaifenxi > 0) {
                this.jtfx_stage_sum(problem_sum["静态分析"]);
                current_table_index += 1;
            }
            else {
                OperationHelper.delete_table(doc, doc_builder, "静态分析SQA", current_table_index);
                this.remove_rows(table, this.current_row_index, 3);
            }
            if(ContentFlags.daimashencha == 0){
                OperationHelper.delete_table(doc, doc_builder, "代码审查SQA", current_table_index);
            }
            else
                current_table_index += 1;
            if(ContentFlags.daimazoucha == 0) {
                OperationHelper.delete_table(doc, doc_builder, "代码走查SQA", current_table_index);
            }
            else
                current_table_index += 1;
            if(ContentFlags.ceshidagang)
                this.xqfx_stage_sum();
            else{
                this.remove_rows(table, current_row_index + 2, 1);
                current_row_index += 2;
            }
            this.csch_stage_sum();
            this.project_sum();
            this.dtcs_stage_sum(problem_sum);
            for(int i = 0; i < this.bookmarks.Count; i ++){
                if(doc_builder.MoveToBookmark(this.bookmarks[i]))
                    doc_builder.Write(this.contents[i]);
            }
            doc.Save(FileConstants.save_root_file);
        }

        public Dictionary<string, int> count_problem(){
            Dictionary<string, int> sum = new Dictionary<string,int>();
            List<List<TestProblem>> problems = ContentFlags.problems.Select(r => r.Value).ToList();
            foreach(List<TestProblem> software_problem in problems){
                foreach(TestProblem p in software_problem) {
                    int temp = 0;
                    for(int i = 0; i < p.qingweiwenti.Count; i ++){
                        temp += p.qingweiwenti[i] + p.yanzhongwenti[i] + p.yibanwenti[i] + p.zhimingwenti[i];
                    }
                    if(sum.ContainsKey(p.cs_stage)){
                        sum[p.cs_stage] += temp;
                    }
                    else
                        sum.Add(p.cs_stage, temp);
                }
            }
            return sum;
        }

        public void wdsc_stage_sum(int num){
            string wdsc_stage_time = "";
            foreach(string k in this.keys){
                if(k.Contains("文档审查")){
                    wdsc_stage_time = ContentFlags.sys_time_dict[k];
                }
            }
            
            string wdsc_stage_bookmark = "文档审查阶段时间";
            contents.Add(wdsc_stage_time);
            bookmarks.Add(wdsc_stage_bookmark);
            string wdsc_summary = "";
            int prob_num = num;
            wdsc_summary += "文档审查共发现" + prob_num.ToString() + "个问题，";
            string bgd_num = prob_num.ToString();
            wdsc_summary += "（形成" + bgd_num + "张问题报告单），问题均已修改";
            contents.Add(wdsc_summary);
            bookmarks.Add("文档审查总结");
            this.current_row_index += 3;
        }

        //删除没有的表格行
        public void remove_rows(Table table, int row_index, int lines){
            for(int i = 0; i < lines; i ++){
                table.Rows[row_index].Remove();
            }
        }

        //静态分析阶段数据
        public void jtfx_stage_sum(int count){
            string jtfx_stage_time = "";
            foreach(string k in this.keys) {
                if(k.Contains("静态分析")) {
                    jtfx_stage_time = ContentFlags.sys_time_dict[k];
                }
            }
            string jtfx_stage_bookmark = "静态分析阶段时间";
            this.contents.Add(jtfx_stage_time);
            this.bookmarks.Add(jtfx_stage_bookmark);
            this.contents.Add(count.ToString());
            this.bookmarks.Add("静态分析问题个数");
            this.current_row_index += 3;
        }

        //需求分析阶段
        public void xqfx_stage_sum() {
            string csdg_stage_time = ContentFlags.sys_time_dict["需求分析阶段"];
            string csdg_stage_bookmark = "需求分析阶段时间";
            this.contents.Add(csdg_stage_time);
            this.bookmarks.Add(csdg_stage_bookmark);
            this.current_row_index += 3;
        }

        //测试策划阶段
        public void csch_stage_sum(){
            string csch_time = ContentFlags.sys_time_dict["策划阶段"];
            string csch_stage_bookmark = "策划阶段时间";
            this.contents.Add(csch_time);
            this.bookmarks.Add(csch_stage_bookmark);
            string text = "";
            if(ContentFlags.wendangshencha > 0)
                text += "文档审查记录、";
            if(ContentFlags.jingtaifenxi > 0)
                text += "静态分析记录、";
            if(ContentFlags.daimashencha > 0)
                text += "代码审查记录、";
            if(ContentFlags.daimazoucha > 0)
                text += "代码走查记录、";
            text += "测试用例执行记录、问题报告单";
            if(ContentFlags.luojiceshi > 0)
                text += "、逻辑测试记录）、逻辑测试结果记录";
            else
                text += ")";
            this.contents.Add(text);
            this.bookmarks.Add("测试过程执行记录");
        }

        //设计、执行阶段
        public void project_sum(){
            string cssj_time = ContentFlags.sys_time_dict["设计阶段"];
            string cssj_stage_bookmark = "设计阶段时间";
            this.contents.Add(cssj_time);
            this.bookmarks.Add(cssj_stage_bookmark);
            List<List<TestExample>> example_list = ContentFlags.example_dict.Select(r => r.Value).ToList();
            List<string> software_name = ContentFlags.example_dict.Select(r => r.Key).ToList();
            int index = 0;
            int count_sum = 0;
            int count_zx_sum = 0;
            string text = "";
            foreach(List<TestExample> list in example_list){
                Dictionary<string, int> sum = new Dictionary<string,int>();
                string temp = "";
                int count_temp = 0;
                int count_zx_temp = 0;
                foreach(TestExample t in list){
                    string name = t.yl_type;
                    int count = int.Parse(t.sjyl_num);
                    sum.Add(name, count);
                    count_temp += count;
                    count_zx_temp += int.Parse(t.zxyl_num);
                    temp += name + count.ToString() + "个测试用例、";
                }
                temp = temp.Substring(0, temp.Length - 1);
                temp = software_name[index] + "共设计了" + count_temp.ToString() + "个测试用例, 其中" + temp +
                    ";\n";
                text += temp;
                count_sum += count_temp;
                count_zx_sum += count_zx_temp;
                index += 1;
            }
            this.contents.Add(text);
            this.bookmarks.Add("测试用例内容");
            this.contents.Add(count_sum.ToString());
            this.bookmarks.Add("测试用例个数");
            this.contents.Add(count_zx_sum.ToString());
            this.bookmarks.Add("执行测试用例个数");
            float rate = (float)count_zx_sum/ContentFlags.code_line/100;
            rate = (float)Math.Round(rate, 2);
            this.contents.Add(rate.ToString());
            this.bookmarks.Add("测试工作质量优良率");
        }

       //动态测试执行阶段与总结阶段
        public void dtcs_stage_sum(Dictionary<string, int> problem_sum) {
            //动态测试执行阶段
            string dtcs_stage_time = "";
            foreach(string k in this.keys) {
                if(k.Contains("动态测试")) {
                    dtcs_stage_time += "配置项首轮测试:" + ContentFlags.sys_time_dict[k] + '\n';
                    break;
                }
            }
            foreach(string k in this.keys) {
                if(k.Contains("动态回归")) {
                    dtcs_stage_time += "配置项回归测试:" + ContentFlags.sys_time_dict[k] + '\n';
                    break;
                }
            }
            foreach(string k in this.keys) {
                if(k.Contains("逻辑")) {
                    dtcs_stage_time += "逻辑测试:" + ContentFlags.sys_time_dict[k] + '\n';
                    break;
                }
            }
            if(ContentFlags.sys_time_dict.ContainsKey("系统测试"))
                dtcs_stage_time += "系统测试:" + ContentFlags.sys_time_dict["系统测试"] + '\n';
            this.contents.Add(dtcs_stage_time);
            this.bookmarks.Add("执行阶段时间");
            //总结阶段开始
            this.contents.Add(ContentFlags.sys_time_dict["总结阶段"]);
            this.bookmarks.Add("总结阶段时间");
            string temp = "";
            int problem_count = 0;
            if(problem_sum.ContainsKey("文档审查")){
                temp += "文档审查共" + problem_sum["文档审查"].ToString() + "个问题，产生问题报告单" +
                    problem_sum["文档审查"].ToString() + "张；";
                problem_count += problem_sum["文档审查"];
            }
            if(problem_sum.ContainsKey("静态分析")) {
                temp += "静态分析共" + problem_sum["静态分析"].ToString() + "个问题，产生问题报告单" +
                    problem_sum["静态分析"].ToString() + "张；";
                problem_count += problem_sum["静态分析"];
            }
            if(problem_sum.ContainsKey("动态测试")) {
                temp += "动态测试共" + problem_sum["动态测试"].ToString() + "个问题，产生问题报告单" +
                    problem_sum["动态测试"].ToString() + "张；";
                problem_count += problem_sum["动态测试"];
            }
            temp = temp.Substring(0, temp.Length - 1);
            string temp1 = "配置项测试共发现问题" + problem_count.ToString() + "个，形成问题报告单" +
                   problem_count.ToString() + "张，";
            temp = temp1 + temp + '，';
            this.contents.Add(temp);
            this.bookmarks.Add("配置项测试总结");
            
        }
    }
}
