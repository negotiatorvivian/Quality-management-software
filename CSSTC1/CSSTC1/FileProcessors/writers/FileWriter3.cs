using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using System.Windows.Forms;
using Aspose.Words.Tables;
using CSSTC1.CommonUtils;

namespace CSSTC1.FileProcessors.writers {
    class FileWriter3 {
        public string questions = "";
        public void write_charts(List<QestionReport> xqns_reports, List<QestionReport> chns_reports, 
            List<QestionReport> ws_report) {
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            bool res = this.write_chart(doc, doc_builder, xqns_reports, InsertionPos.csxqns_section, 
                InsertionPos.csxqns_sec_table1, InsertionPos.csxqns_ques_row);
            if(res)
                this.write_wtzz_chart(doc, doc_builder, xqns_reports, InsertionPos.csxqns_section,
                InsertionPos.csxqns_sec_table2, InsertionPos.csxqns_solu_row);
            bool res1 = this.write_chart(doc, doc_builder, chns_reports, InsertionPos.cschns_section,
                InsertionPos.cschns_sec_table1, InsertionPos.cschns_ques_row);
            if(res1)
                this.write_wtzz_chart(doc, doc_builder, chns_reports, InsertionPos.cschns_section,
                InsertionPos.cschns_sec_table2, InsertionPos.cschns_solu_row);
            bool res2 = this.write_chart(doc, doc_builder, ws_report, InsertionPos.csxqws_section,
                InsertionPos.csxqws_sec_table1, InsertionPos.csxqws_ques_row);
            if(res2)
                this.write_wtzz_chart(doc, doc_builder, ws_report, InsertionPos.csxqws_section,
                InsertionPos.csxqws_sec_table2, InsertionPos.csxqws_solu_row);
            this.write_pzgg_chart(doc, doc_builder, ws_report, "测试说明与计划更改内容");
            doc.Save(FileConstants.save_root_file);

            MessageBox.Show("需求规格与说明文件读取完成");
        }

        //偏差与问题报告
        public bool write_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files, 
            int sec_index, int sec_table_index, int row_index) {
            int section_index = sec_index + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
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
            int section_index = sec_index + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
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
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(solutions);
        }

        //会议签到表
        public void write_hyqdb_chart(Document doc, DocumentBuilder doc_builder, List<string> names) {
            int section_index = InsertionPos.csxqqd_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            int sec_table_index = InsertionPos.csxqqd_sec_table;
            OperationHelper.fill_pszcy_info(doc_builder, names);
            bool res = OperationHelper.conference_signing(doc, doc_builder, section_index, sec_table_index);
            if(!res)
                MessageBox.Show("请先填写测试需求分析与策划阶段的信息");
            else{
                int section_index1 = InsertionPos.cschqd_section + ContentFlags.lingqucishu * 2 + 
                    ContentFlags.pianli_1;
                int sec_table_index1 = InsertionPos.cschqd_sec_table;
                OperationHelper.conference_signing(doc, doc_builder, section_index1, sec_table_index1);
                doc.Save(FileConstants.save_root_file);
            }
        }
    }
}
