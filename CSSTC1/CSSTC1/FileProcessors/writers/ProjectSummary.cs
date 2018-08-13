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
using CSSTC1.FileProcessors.writers.ConfigurationItemTest;

namespace CSSTC1.FileProcessors.writers {
    class ProjectSummary {
        private List<int> time_diff = new List<int>();
        private string pzcs_record = "";
        private string pzcs_record1 = "";
        private string pzcs_record2 = "";
        private string software_name = "";
        public ProjectSummary(){
            time_diff.Add(ContentFlags.lingqucishu * 2);
            time_diff.Add(ContentFlags.pianli_1);
            time_diff.Add(ContentFlags.pianli_2);
            time_diff.Add(ContentFlags.wendangshencha);
            time_diff.Add(ContentFlags.jingtaifenxi);
            time_diff.Add(ContentFlags.daimashencha);
            time_diff.Add(ContentFlags.daimazoucha);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.pianli_3);
            time_diff.Add(ContentFlags.ceshijiuxu2);
            time_diff.Add(ContentFlags.beiceruanjianshuliang1 * 2);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.peizhiceshi);
            time_diff.Add(ContentFlags.luojiceshi);
            time_diff.Add(ContentFlags.xitongceshi);
            time_diff.Add(ContentFlags.xitonghuiguiceshi);
        }
        public void write_charts(List<QestionReport> ns_report, List<QestionReport> ws_report) {
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            string software_name = doc.Range.Bookmarks["软件名称"].Text;
            this.software_name = software_name;
            this.write_time_line(doc, doc_builder);
            OperationHelper.conference_signing(doc, doc_builder, "鉴定测评报告时间", 
                InsertionPos.xtzj_hyqdb_sec_table);
            //内审意见表
            this.write_pcywt_chart(doc, doc_builder, ns_report, InsertionPos.xtzj_jdcp_section,
                InsertionPos.xtzj_wtbg_sec_table, this.time_diff);
            this.write_pcywtzz_chart(doc, doc_builder, ns_report,InsertionPos.xtzj_jdcp_section,
                InsertionPos.xtzj_wtzz_sec_table, this.time_diff);
            //外审意见表
            this.write_pcywt_chart(doc, doc_builder, ws_report, InsertionPos.xtzj_jdcp_section1,
                InsertionPos.xtzj_wtbg_sec_table, this.time_diff);
            this.write_pcywtzz_chart(doc, doc_builder, ws_report, InsertionPos.xtzj_jdcp_section1,
                InsertionPos.xtzj_wtzz_sec_table, this.time_diff);
            this.write_pzgg_chart(doc, doc_builder, ws_report, "鉴定测评配置更改内容");
            //第二个出库申请单
            int row_count = this.write_cksqd_chart2(doc, doc_builder, InsertionPos.xtzj_cksqd_section1,
                InsertionPos.xtzj_cksqd_sec_table, InsertionPos.sj_rksqd_name_row,
                InsertionPos.xtzj_cssmfj_row_index2, this.time_diff);
            //int temp = row_count + InsertionPos.xtzj_wdscjl_row_index2;
            //第一个出库申请单和入库申请单
            List<int> row_counts = this.write_cksqd_chart(doc, doc_builder, InsertionPos.xtzj_cksqd_section, 
                InsertionPos.xtzj_cksqd_section1, InsertionPos.xtzj_cksqd_sec_table,
                InsertionPos.sj_rksqd_name_row, InsertionPos.xtzj_wdscjl_row_index1, row_count,
                this.time_diff);
            if(row_counts != null){
                int start_row = this.write_cksqd_chart1(doc, doc_builder, InsertionPos.xtzj_cksqd_section, 
                    InsertionPos.xtzj_cksqd_section1, InsertionPos.xtzj_cksqd_sec_table, 
                    InsertionPos.sj_rksqd_name_row, row_counts[0], row_counts[1], this.time_diff);
                this.write_cksqd_rkwd_chart(doc, doc_builder, InsertionPos.xtzj_cksqd_section1,
                     InsertionPos.xtzj_cksqd_sec_table, InsertionPos.sj_rksqd_name_row,
                     InsertionPos.sj_rksqd_iden_row, start_row, this.time_diff);
            }
            doc.Save(FileConstants.save_root_file);

        }

        //阶段主要时间
        public void write_time_line(Document doc, DocumentBuilder doc_builder){
            List<string> times = new List<string>();
            List<string> bookmarks = new List<string>();
            string jdcpps_time = TimeStamp.csbgps_time;
            times.Add(TimeStamp.csbgps_format_time);
            bookmarks.Add("鉴定测评报告外审时间");
            string jdcpns_time = DateHelper.cal_time(jdcpps_time, 7);
            times.Add(jdcpns_time);
            bookmarks.Add("鉴定测评报告时间");
            string lxwtf_time = DateHelper.cal_time(jdcpps_time, 3);
            times.Add(lxwtf_time);
            bookmarks.Add("联系委托方第十二次");
            string jdcpck_time = DateHelper.cal_time(jdcpps_time, -3);
            times.Add(jdcpck_time);
            bookmarks.Add("鉴定测评报告出库时间");
            string lxwtf_time1 = DateHelper.cal_time(jdcpps_time, -10);
            times.Add(lxwtf_time1);
            bookmarks.Add("联系委托方第十三次");
            int i = 0;
            foreach(string mark in bookmarks){
                if(doc_builder.MoveToBookmark(mark)){
                    doc_builder.Write(times[i]);
                    i += 1;
                }
            }
        }

        //偏差与问题报告
        public bool write_pcywt_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files, 
            int section_index, int sec_table_index, List<int> time_diff) {
            foreach(int i in time_diff)
                section_index += i;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                //Row row = table.Rows[index_row];
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.csdgns_ques_row, 0);
                doc_builder.Write(report.wenti);
                index_row += 1;
            }
            return true;
        }

        //偏差与问题追踪
        public bool write_pcywtzz_chart(Document doc, DocumentBuilder doc_builder,List<QestionReport> files, 
            int section_index, int sec_table_index, List<int> time_diff) {
            foreach(int i in time_diff)
                section_index += i;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.csdgns_solu_row, 0);
                doc_builder.Write(report.chulicuoshi);
                index_row += 1;
            }
            return true;
        }

       //配置状态更改表
        public void write_pzgg_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files,
            string bookmark) {
            string solutions = "";
            foreach(QestionReport report in files) {
                string temp = report.chulicuoshi;
                if(temp.Contains("已采纳"))
                    temp = temp.Substring(5, temp.Length - 5);
                solutions += temp + '\n';
            }
            solutions = solutions.Substring(0, solutions.Length - 1);
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(solutions);
        }

        //文档审查等模块记录
        public List<int> write_cksqd_chart(Document doc, DocumentBuilder doc_builder, int section, int section1,
            int sec_table_index, int name_row_index, int row_index, int row_index1, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
                section1 += i;
            }
            List<string> contents = new List<string>();
            string pzcs_content = "";
            if(ContentFlags.wendangshencha > 0){
                string text = this.software_name + "文档审查结果及记录";
                contents.Add(text);
                pzcs_content += text + " V1.1" + '\n';
            }
            if(ContentFlags.jingtaifenxi > 0) {
                string text = this.software_name + "静态分析结果及记录";
                contents.Add(text);
                pzcs_content += text + " V1.1" + '\n';
            }
            if(ContentFlags.daimashencha > 0) {
                string text = this.software_name + "代码审查结果及记录";
                contents.Add(text);
                pzcs_content += text + " V1.1" + '\n';
            }
            if(ContentFlags.daimazoucha > 0) {
                string text = this.software_name + "代码走查结果及记录";
                contents.Add(text);
                pzcs_content += text + " V1.1" + '\n';
            }
            pzcs_content += this.software_name + "测试用例执行记录" + " V1.2" + '\n';
            this.pzcs_record = pzcs_content;
            if(contents.Count == 0)
                return null;
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            Table table1 = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index + 1, true);
            doc_builder.MoveToSection(section1);
            Table table2 = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string temp in contents) {
                doc_builder.MoveToSection(section);
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                var row1 = table1.Rows[row_index].Clone(true);
                table1.Rows.Insert(row_index + 1, row1);
                var row2 = table2.Rows[row_index1].Clone(true);
                table2.Rows.Insert(row_index1 + 1, row2);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                doc_builder.MoveToCell(sec_table_index + 1, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                doc_builder.MoveToSection(section1);
                doc_builder.MoveToCell(sec_table_index, row_index1, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
                row_index1 += 1;
            }
            table.Rows[row_index].Remove();
            table1.Rows[row_index].Remove();
            table2.Rows[row_index1].Remove();
            List<int> indexes = new List<int>();
            indexes.Add(row_index + 1);
            indexes.Add(row_index1 + 1);
            return indexes;
        }


        //配置项动态测试原始记录+配置项动态回归测试用例集
        public int write_cksqd_chart1(Document doc, DocumentBuilder doc_builder, int section, int section1,
            int sec_table_index, int name_row_index, int row_index, int row_index1, 
            List<int> time_diff) {
            int origin_row_index = row_index;
            foreach(int i in time_diff) {
                section += i;
                section1 += i;
            }
            List<string> contents = new List<string>();
            //List<string> contents_new = new List<string>();
            int count = 1;
            string content = "";
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = this.software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态测试原始记录";
                count += 1;
                contents.Add(text);
                content += text + "V1.0" + '\n';
            }
            if(ContentFlags.xitongceshi > 0){
                string xtrj_text = this.software_name + "测试执行记录附件" + count.ToString() + "-" +
                this.software_name + "系统测试原始记录";
                contents.Add(xtrj_text);
                content += xtrj_text + " V1.0" + '\n';
            }
            count = 1;
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = this.software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态回归测试原始记录";
                count += 1;
                contents.Add(text);
                content += text + "V1.0" + '\n';
            }
            //contents.Concat(contents_new);
            this.pzcs_record += content;
            this.pzcs_record = this.pzcs_record.Substring(0, this.pzcs_record.Length - 1);
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            Table table1 = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index + 1, true);
            doc_builder.MoveToSection(section1);
            Table table2 = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            
            doc_builder.MoveToSection(section);
            foreach(string temp in contents) {
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);    
                row_index += 1;
            }
            table.Rows[row_index].Remove();
            
            row_index = origin_row_index;
            foreach(string temp in contents){
                var row1 = table1.Rows[row_index].Clone(true);
                table1.Rows.Insert(row_index + 1, row1);
                doc_builder.MoveToCell(sec_table_index + 1, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            table1.Rows[row_index].Remove();

            doc_builder.MoveToSection(section1);
            foreach(string temp in contents) {
                var row2 = table2.Rows[row_index1].Clone(true);
                table2.Rows.Insert(row_index1 + 1, row2);
                doc_builder.MoveToCell(sec_table_index, row_index1, name_row_index, 0);
                doc_builder.Write(temp);
                row_index1 += 1;
            }
            table2.Rows[row_index1].Remove();

            if(ContentFlags.luojiceshi == 0) {
                table.Rows[row_index].Remove();
                table1.Rows[row_index].Remove();
                table2.Rows[row_index1].Remove();
            }
            if(doc_builder.MoveToBookmark("鉴定测评配置状态报告1"))
                doc_builder.Write(this.pzcs_record);
            return row_index1 + 1;
        }

        //最后一张出库申请表测试用例集部分
        public int write_cksqd_chart2(Document doc, DocumentBuilder doc_builder, int section,
            int sec_table_index, int name_row_index, int row_index, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
            }
            
            List<string> contents = new List<string>();
            List<string> contents_new = new List<string>();
            int count = 1;
            string content = "";
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = this.software_name;
                text += "测试说明附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态测试用例集";
                count += 1;
                contents.Add(text);
                content += text + "V1.0" + '\n';
            }
            if(ContentFlags.xitongceshi > 0) {
                string xtrj_text = this.software_name + "测试说明附件" + count.ToString() + "-" +
                this.software_name + "系统测试用例集";
                contents.Add(xtrj_text);
                count += 1;
                content += xtrj_text + " V1.0" + '\n';
            }
            //count = 1;
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = this.software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态回归测试用例集";
                count += 1;
                contents_new.Add(text);
                content += text + "V1.0" + '\n';
            }
            this.pzcs_record1 += content;
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string temp in contents) {
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            foreach(string temp in contents_new) {
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            table.Rows[row_index].Remove();
            if(doc_builder.MoveToBookmark("鉴定测评配置状态报告2"))
                doc_builder.Write(this.pzcs_record1);
            return row_index + 1;
        }

        //第二阶段和第五阶段入库的被测件文档
        public void write_cksqd_rkwd_chart(Document doc, DocumentBuilder doc_builder, int section,
            int sec_table_index, int name_row_index, int iden_row_index, int row_index, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
            }
            string content = "";
            Dictionary<string, string> rukuwendang_dict = ContentFlags.rukuwendang_dict;
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            Regex reg = new Regex("[v|V][0-9][.][0-9]+");
            foreach(string name in rukuwendang_dict.Keys) {
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(rukuwendang_dict[name]);
                row_index += 1;
                MatchCollection matches = reg.Matches(rukuwendang_dict[name], 0);
                if(matches == null)
                    continue;
                string version = matches[matches.Count - 1].Value;
                string text = name + ' ' + version + '\n';
                content += text;
            }
            table.Rows[row_index].Remove();
            this.pzcs_record2 += content;
            //被测件清单部分
            int bcjqd_count = 0;
            bcjqd_count = ContentFlags.lingqucishu + ContentFlags.wendangshencha + ContentFlags.jingtaifenxi / 15 
                * 2 + ContentFlags.daimashencha / 13 * 2 + ContentFlags.daimazoucha / 13 * 2 + 
                ContentFlags.ceshijiuxu2 / 3 * 2 + ContentFlags.peizhiceshi /6 + ContentFlags.xitongceshi/5 + 
                ContentFlags.xitonghuiguiceshi/5;
            for(int i = 0; i < bcjqd_count; i ++){
                var row = table.Rows[row_index].Clone(true);
                table.Rows.Insert(row_index + 1, row);
                row_index += 1;
            }
            table.Rows[row_index].Remove();
            if(doc_builder.MoveToBookmark("鉴定测评配置状态报告3"))
                doc_builder.Write(this.pzcs_record2);
        }


    }
}
