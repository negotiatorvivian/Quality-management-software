﻿using System;
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

namespace CSSTC1.FileProcessors.writers.CodeReview_Walkthrough {
    class FileWriter7 {
        private List<int> time_diff = new List<int>();
        private Table hardware_table;
        private string type;
        private string[] dmsc_bookmarks = { "代码审查软件名称", "代码审查被测软件", "代码审查回归被测软件" };
        private string[] dmzc_bookmarks = { "代码走查被测软件", "代码走查被测软件配置报告", 
                                              "代码走查被测软件配置报告1" };
        public FileWriter7(string type, Table table){
            time_diff.Add(ContentFlags.lingqucishu * 2);
            time_diff.Add(ContentFlags.pianli_1);
            time_diff.Add(ContentFlags.pianli_2);
            time_diff.Add(ContentFlags.wendangshencha);
            time_diff.Add(ContentFlags.jingtaifenxi);
            this.type = type;
            string[] marks = this.dmsc_bookmarks;
            if(type.Equals("代码走查")){
                time_diff.Add(ContentFlags.daimashencha);
                marks = this.dmzc_bookmarks;
            }
            this.hardware_table = table;
            this.write_charts(marks);
        }

        public void write_charts(string[] bookmarks){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            if(this.type.Equals("代码审查"))
                this.write_time_line(doc, doc_builder);
            else
                this.write_dmzc_time_line(doc, doc_builder);
            List<StaticAnalysisFile> code_review_list = new List<StaticAnalysisFile>();
            switch(this.type){
                case "代码审查":
                    code_review_list = ContentFlags.dmsc_software_info;
                    break;
                case "代码走查":
                    code_review_list = ContentFlags.dmzc_software_info;
                    break;
                default:
                    code_review_list = ContentFlags.dmsc_software_info;
                    break;
            }

            this.write_file_charts(doc, doc_builder, code_review_list, bookmarks[0], bookmarks[1]);
            //测试环境核查单
            ChartHelper.append_content(doc, doc_builder, this.hardware_table, InsertionPos.dmsc_chhjhcd_section,
                InsertionPos.dmsc_cshj_sec_table1, 2, this.time_diff);
            this.write_cehjhcd_software(doc, doc_builder, code_review_list, 
                InsertionPos.dmsc_chhjhcd_section, InsertionPos.dmsc_rjhj_sec_table1, 1 ,this.time_diff);
            this.write_file_charts_2(doc, doc_builder, code_review_list, bookmarks[2]);
            doc.Save(FileConstants.save_root_file);
            //MessageBox.Show(this.type + "填写结束");
        }

        //填写代码审查时间线
        public void write_time_line(Document doc, DocumentBuilder doc_builder){
            List<string> bookmarks = new List<string>();
            List<string> values = new List<string>();
            if(!ContentFlags.dmsc_same){
                string dmsc_time = TimeStamp.dmsc_format_time;
                bookmarks.Add("代码审查时间");
                values.Add(dmsc_time);
                DateTime t1 = DateHelper.cal_date(TimeStamp.dmsc_time, 0);
                ContentFlags.time_dict2.Add("代码审查时间", t1);
            }
            string dmsc_lxwtf_time = DateHelper.cal_time(TimeStamp.dmsc_time, 1);
            bookmarks.Add("代码审查联系委托方");
            values.Add(dmsc_lxwtf_time);
            DateTime t2 = DateHelper.cal_date(TimeStamp.dmsc_time, 1);
            ContentFlags.time_dict1.Add("代码审查联系委托方", t2);

            string dmschg_time = TimeStamp.dmschg_format_time;
            bookmarks.Add("代码审查回归时间");
            values.Add(dmschg_time);
            DateTime t3 = DateHelper.cal_date(TimeStamp.dmschg_time, 0);
            ContentFlags.time_dict2.Add("代码审查回归时间", t3);
            string dmscqr_time = TimeStamp.dmscqr_format_time;
            bookmarks.Add("代码审查确认时间");
            values.Add(dmscqr_time);
            string dmsc_rk_time = DateHelper.cal_time(TimeStamp.dmscqr_time, 1);
            bookmarks.Add("代码审查文档入库时间");
            values.Add(dmsc_rk_time);
            DateTime t4 = DateHelper.cal_date(TimeStamp.dmscqr_time, 1);
            ContentFlags.time_dict1.Add("代码审查文档入库时间", t4);
            ContentFlags.time_dict2.Add("代码审查文档入库时间", t4);
           
            int count = 0;
            foreach(string bookmark in bookmarks){
                if(doc_builder.MoveToBookmark(bookmark))
                    doc_builder.Write(values[count]);
                count += 1;
            }
        }

        //填写代码走查时间线
        public void write_dmzc_time_line(Document doc, DocumentBuilder doc_builder) {
            List<string> bookmarks = new List<string>();
            List<string> values = new List<string>();
            //if(!ContentFlags.dmzc_same){
                string dmzc_time = TimeStamp.dmzc_format_time;
                bookmarks.Add("代码走查时间");
                values.Add(dmzc_time);
                DateTime t1 = DateHelper.cal_date(TimeStamp.dmzc_time, 0);
                ContentFlags.time_dict2.Add("代码走查时间", t1);
           // }
            string dmzc_lxwtf_time = DateHelper.cal_time(TimeStamp.dmzc_time, 1);
            bookmarks.Add("代码走查联系委托方");
            DateTime t2 = DateHelper.cal_date(TimeStamp.dmzc_time, 1);
            ContentFlags.time_dict1.Add("代码走查联系委托方", t2);
            values.Add(dmzc_lxwtf_time);
            string dmzchg_time = TimeStamp.dmzchg_format_time;
            bookmarks.Add("代码走查回归时间");
            values.Add(dmzchg_time);
            DateTime t3 = DateHelper.cal_date(TimeStamp.dmzchg_time, 0);
            ContentFlags.time_dict2.Add("代码走查回归时间", t3);
            string dmzcqr_time = TimeStamp.dmzcqr_format_time;
            bookmarks.Add("代码走查确认时间");
            values.Add(dmzcqr_time);
            //DateTime t4 = DateHelper.cal_date(TimeStamp.dmzcqr_time, 0);
            //ContentFlags.time_dict2.Add("代码走查确认时间", (int)t4.Ticks);
            string dmzc_rk_time = DateHelper.cal_time(TimeStamp.dmzcqr_time, 1);
            bookmarks.Add("代码走查入库申请时间");
            values.Add(dmzc_rk_time);
            DateTime t5 = DateHelper.cal_date(TimeStamp.dmzcqr_time, 1);
            ContentFlags.time_dict1.Add("代码走查入库申请时间", t5);
            ContentFlags.time_dict2.Add("代码走查入库申请时间", t5);
            int count = 0;
            foreach(string bookmark in bookmarks) {
                if(doc_builder.MoveToBookmark(bookmark))
                    doc_builder.Write(values[count]);
                count += 1;
            }
        }

        public bool write_file_charts(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> code_review_list, string bookmark1, string bookmark2) {
            if(code_review_list.Count == 0)
                return false;
            this.write_lxwtf_chart(doc, doc_builder, code_review_list, bookmark1);
            Dictionary<string, StaticAnalysisFile> dict = this.write_bcjqd_chart(doc, doc_builder, 
                code_review_list);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section,
                InsertionPos.dmsc_bcjdbd_sec_table1, 1, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            ChartHelper.write_rksqd_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section,
                InsertionPos.dmsc_bcjlqqd_sec_table1, 2, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_code_row_index, InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            this.update_code_language(dict);
            ChartHelper.write_rksqd_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section,
                InsertionPos.dmsc_rksqd_sec_table1, 3, InsertionPos.sj_rksqd_name_row, 
                InsertionPos.sj_code_row_index, InsertionPos.sj_rksqd_iden_row, this.time_diff);
            this.write_pzbbd_chart(doc, doc_builder, code_review_list, bookmark2);
            
            return true;
        }

        //同步编程语言信息
        public void update_code_language(Dictionary<string, StaticAnalysisFile> dict) {
            foreach(string name in dict.Keys){
                for(int i = 0; i< ContentFlags.static_files.Count; i++){
                    if(dict[name].rj_mingcheng.Equals(ContentFlags.static_files[i].rj_mingcheng)){
                        dict[name].bc_yuyan = ContentFlags.static_files[i].bc_yuyan;
                        break;
                    }
                }
                if(dict[name].bc_yuyan.Length == 0)
                    dict[name].bc_yuyan = "C/C++";
            }
        }


        //填写第一个联系委托方单
        public void write_lxwtf_chart(Document doc, DocumentBuilder doc_builder, 
            List<StaticAnalysisFile> code_review_list, string mark){
            string text = "";
            foreach(StaticAnalysisFile file in code_review_list){
                text += file.rj_mingcheng + file.jtfx_fanwei + '、';
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写第一个配置报告单
        public void write_pzbbd_chart(Document doc, DocumentBuilder doc_builder,

            List<StaticAnalysisFile> code_review_list, string bookmark) {
            string text = "";
            foreach(StaticAnalysisFile file in code_review_list) {
                text += file.rj_mingcheng + file.jtfx_fanwei + '\t' + file.xt_banben + '\n';
            }
            //text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(text);
        }

        //填写第二个配置报告单(回归版本)
        public void write_pzbbd_chart_2(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> code_review_list, string mark) {
            string text = "";
            foreach(StaticAnalysisFile file in code_review_list) {
                text += file.rj_mingcheng + file.jtfx_fanwei + '\t' + file.hg_banben + '\n';
            }
            //text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写被测件清单
        public Dictionary<string, StaticAnalysisFile> write_bcjqd_chart(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> code_review_list) {
            Dictionary<string, StaticAnalysisFile> dict = this.set_file_id(doc, doc_builder, code_review_list, 
                false);
            ChartHelper.write_bcjqd2_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section,
                InsertionPos.dmsc_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row, 
                this.time_diff);
            return dict;
        }

        //填写文件标识
        public Dictionary<string, StaticAnalysisFile> set_file_id(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> software_items, bool update) {
            Dictionary<string, StaticAnalysisFile> dict = new Dictionary<string, StaticAnalysisFile>();
            string id = "";
            string year = "";
            if(doc_builder.MoveToBookmark("项目标识"))
                id = doc.Range.Bookmarks["项目标识"].Text;
            else
                return null;
            if(doc_builder.MoveToBookmark("年份"))
                year = doc.Range.Bookmarks["年份"].Text;
            else
                return null;
            if(ContentFlags.software_dict.Count == 0){
                int count = 1;
                foreach(StaticAnalysisFile item in software_items) {
                    string version = "";
                    if(!update)
                        version = item.xt_banben;
                    else
                        version = item.hg_banben;
                    string key = NamingRules.pre_name;
                    key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C19";
                    if(count < 10) {
                        key += "(0" + count.ToString() + ')' + "-" + version + '-' + year;
                    }
                    else
                        key += '(' + count.ToString() + ')' + "-" + version + '-' + year;
                    dict.Add(key, item);
                    count += 1;
                }
            }
            else{
                int count = ContentFlags.software_dict.Count + 1;
                foreach(StaticAnalysisFile item in software_items) {
                    string key = NamingRules.pre_name;
                    string version = "";
                    if(!update)
                        version = item.xt_banben;
                    else
                        version = item.hg_banben;
                    key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C19";
                    if(count < 10) {
                        key += "(0" + count.ToString() + ')' + "-" + version + '-' + year;
                    }
                    else
                        key += '(' + count.ToString() + ')' + "-" + version + '-' + year;
                    dict.Add(key, item);
                    count += 1;
                }
            }
            if(update){
                foreach(string key in dict.Keys){
                    ContentFlags.software_dict.Add(key, dict[key]);
                }
            }
            return dict;
        }

        //填写测试环境核查单软件环境
        public void write_cehjhcd_software(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> software_items, int section_index, int sec_table_index, int
            row_index, List<int> time_diff) {
            int flag = row_index;
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(StaticAnalysisFile file in software_items) {
                if(row_index < software_items.Count + flag - 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.dmsc_rjhj_name_row, 0);
                string name = file.rj_mingcheng + file.jtfx_fanwei;
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.dmsc_rjhj_version_row, 0);
                doc_builder.Write(file.xt_banben);
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.dmsc_rjhj_orig_row, 0);
                doc_builder.Write(file.yz_danwei);
                row_index += 1;
            }
        }
        
        //第二次文档调拨与入库
        public bool write_file_charts_2(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> code_review_list, string mark) {
            this.write_pzbbd_chart_2(doc, doc_builder, code_review_list, mark);
            Dictionary<string, StaticAnalysisFile> dict = this.write_bcjqd_chart_2(doc, doc_builder,
                code_review_list);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section1,
                InsertionPos.dmsc_bcjdbd_sec_table1, 1, InsertionPos.sj_bcjdbd_name_row,
                InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            ChartHelper.write_rksqd_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section1,
                InsertionPos.dmsc_bcjlqqd_sec_table1, 2, InsertionPos.sj_bcjdbd_name_row, 
                ContentFlags.missing, InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            ChartHelper.write_rksqd_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section1,
                InsertionPos.dmsc_rksqd_sec_table1, 3, InsertionPos.sj_rksqd_name_row, 
                InsertionPos.sj_code_row_index, InsertionPos.sj_rksqd_iden_row, this.time_diff);
            return true;
        }

        //填写第二个被测件清单p152
        public Dictionary<string, StaticAnalysisFile> write_bcjqd_chart_2(Document doc, DocumentBuilder 
            doc_builder, List<StaticAnalysisFile> code_review_list) {
            Dictionary<string, StaticAnalysisFile> dict = this.set_file_id(doc, doc_builder, code_review_list,
                true);
            ChartHelper.write_bcjqd2_chart(doc, doc_builder, dict, InsertionPos.dmsc_bcjqd_section1,
                InsertionPos.dmsc_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row,
                this.time_diff);
            return dict;
        }
    }
}
