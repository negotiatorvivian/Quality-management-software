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

namespace CSSTC1.FileProcessors.writers.part3_4 {
    //第五章 三四节
    class FileWriter5 {
        public string questions = "";
        List<int> times = new List<int>();
        Dictionary<string, StaticAnalysisFile> software_dict;
        public FileWriter5(List<QestionReport> cssmns_reports) {
            this.software_dict = ContentFlags.software_dict;
            times.Add(ContentFlags.pianli_1);
            times.Add(ContentFlags.pianli_2);
            times.Add(ContentFlags.lingqucishu * 2);
            times.Add(ContentFlags.wendangshencha);
            times.Add(ContentFlags.jingtaifenxi);
            times.Add(ContentFlags.daimashencha);
            times.Add(ContentFlags.daimazoucha);
            this.write_charts(cssmns_reports);
            //if(jxwt_reports.Count > 0)
            //    this.write_charts(jxwt_reports);  
        }

        //项目SQA偏差与问题追踪报告
        public void write_charts(List<QestionReport> cssmns_reports) {
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            //测试说明内部评审时间
            if(doc_builder.MoveToBookmark("测试说明内部评审时间")){
                doc_builder.Write(TimeStamp.cssmps_format_time);
                DateTime cssmps_format_time = DateHelper.cal_date(TimeStamp.cssmps_time, 0);
                ContentFlags.time_dict2.Add("测试说明内部评审时间", cssmps_format_time);
            }
            bool res = this.write_chart(doc, doc_builder, cssmns_reports, InsertionPos.cssmns_section,
                InsertionPos.cssmns_sec_table1, InsertionPos.cssmns_ques_row);
            if(res)
                this.write_wtzz_chart(doc, doc_builder, cssmns_reports, InsertionPos.cssmns_section,
                InsertionPos.cssmns_sec_table2, InsertionPos.cssmns_solu_row);
            int section_index = InsertionPos.cssmns_section;
            foreach(int i in times)
                section_index += i;
            OperationHelper.conference_signing(doc, doc_builder, section_index,
                InsertionPos.cssmns_hyqdb_table);
            
            List<string> contents = this.write_pzzt_chart(doc, doc_builder, "测试说明入库清单1");
            this.write_rksqd_chart(doc, doc_builder, contents, InsertionPos.cssmps_rksqd_section,
                InsertionPos.cssmps_rksqd_sec_table, 4, InsertionPos.sj_rksqd_name_row, times);
            doc.Save(FileConstants.save_root_file);


        }

        //项目SQA偏差与问题报告
        public bool write_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files,
            int sec_index, int sec_table_index, int row_index) {
            int section_index = sec_index;
            foreach(int i in times)
                section_index += i;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, row_index, 0);
                doc_builder.Write(report.wenti);
                index_row += 1;
            }
            return true;
        }

        //偏差与问题追踪
        public bool write_wtzz_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files,
            int sec_index, int sec_table_index, int row_index) {
            int section_index = sec_index;
            foreach(int i in times)
                section_index += i;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, row_index, 0);
                doc_builder.Write(report.chulicuoshi);
                index_row += 1;
            }
            return true;
        }


        //会议签到表
        public void write_hyqdb_chart(Document doc, DocumentBuilder doc_builder, List<string> names) {
            int section_index = InsertionPos.csxqqd_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            int sec_table_index = InsertionPos.csxqqd_sec_table;
            OperationHelper.fill_pszcy_info(doc_builder, names);
            bool res = OperationHelper.conference_signing(doc, doc_builder, section_index, sec_table_index);
            if(!res)
                MessageBox.Show("请先填写测试需求分析与策划阶段的信息");
            else {
                int section_index1 = InsertionPos.cschqd_section + ContentFlags.lingqucishu * 2 +
                    ContentFlags.pianli_1;
                int sec_table_index1 = InsertionPos.cschqd_sec_table;
                OperationHelper.conference_signing(doc, doc_builder, section_index1, sec_table_index1);
                doc.Save(FileConstants.save_root_file);
            }
        }

        //配置状态报告
        public List<string> write_pzzt_chart(Document doc, DocumentBuilder doc_builder, string mark) {
            string text = "";
            string default1 = "测试说明附件";
            string default2 = "配置项动态测试用例集";
            List<string> contents = new List<string>();
            if(doc_builder.MoveToBookmark("软件名称")){
                string softawre_name = doc.Range.Bookmarks["软件名称"].Text;
                int count = 1;
                foreach(string key in this.software_dict.Keys){
                    string temp = softawre_name + default1 + count.ToString() + ':';
                    temp += software_dict[key].rj_mingcheng + default2;
                    contents.Add(temp);
                    text +=  temp + '\t' + software_dict[key].xt_banben + '\n';
                    count += 1;
                }
                text = text.Substring(0, text.Length - 1);
                if(doc_builder.MoveToBookmark(mark))
                    doc_builder.Write(text);
            }
            return contents;
        }

        //填写入库申请单
        public void write_rksqd_chart(Document doc, DocumentBuilder doc_builder, List<string> contents,
            int section_index, int sec_table_index, int row_index, int name_row_index, List<int> time_diff) {
            int flag = row_index;
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string key in contents) {
                if(row_index < contents.Count + flag - 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                //doc.Save(FileConstants.save_root_file);
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(key);
                row_index += 1;
            }
        }
    }
}
