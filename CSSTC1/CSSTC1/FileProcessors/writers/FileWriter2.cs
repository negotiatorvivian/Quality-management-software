using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using Aspose.Words.Tables;
using System.Windows.Forms;

namespace CSSTC1.FileProcessors.writers {
    class FileWriter2 {
        public string questions = "";
        public void write_charts(List<QestionReport> ns_report, List<QestionReport> ws_report) {
            Document doc = new Document(FilePaths.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            bool res = this.write_csdgnswt_chart(doc, doc_builder, ns_report);
            if(res)
                this.write_csdgnswtzz_chart(doc, doc_builder, ns_report);
            doc.Save(FilePaths.save_root_file);
            bool res1 = this.write_csdgwswt_chart(doc, doc_builder, ws_report);
            if(res1)
                this.write_csdgwswtzz_chart(doc, doc_builder, ws_report);
            this.write_file_list(doc, doc_builder, InsertionPos.pzgljh_section,
                InsertionPos.pzgljh_sec_table, ContentFlags.beicejianqingdan, ContentFlags.content_list);
            this.write_pzgg_chart(doc, doc_builder, ws_report);
            doc.Save(FilePaths.save_root_file);
            MessageBox.Show("测试大纲文件读取完成");

        }

        //测试大纲内审偏差与问题报告
        public bool write_csdgnswt_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files) {
            int section_index = InsertionPos.csdgns_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, 
                InsertionPos.csdgns_sec_table1, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                //Row row = table.Rows[index_row];
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(InsertionPos.csdgns_sec_table1, index_row, InsertionPos.csdgns_ques_row, 0);
                doc_builder.Write(report.wenti);
                index_row += 1;
               
            }
            //doc.Save(FilePaths.save_root_file);
            return true;
        }

        //测试大纲内审偏差与问题追踪
        public bool write_csdgnswtzz_chart(Document doc, DocumentBuilder doc_builder, 
            List<QestionReport> files) {
            int section_index = InsertionPos.csdgns_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, 
                InsertionPos.csdgns_sec_table2, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(InsertionPos.csdgns_sec_table2, index_row, InsertionPos.csdgns_solu_row, 0);
                doc_builder.Write(report.chulicuoshi);
                index_row += 1;
            }
            return true;
        }

        //测试大纲外审偏差与问题报告
        public bool write_csdgwswt_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files) {
            int section_index = InsertionPos.csdgws_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, 
                InsertionPos.csdgws_sec_table1, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(InsertionPos.csdgws_sec_table1, index_row, InsertionPos.csdgns_ques_row, 0);
                doc_builder.Write(report.wenti);
                index_row += 1;
            }
            return true;
        }

        //测试大纲外审偏差与问题追踪报告
        public bool write_csdgwswtzz_chart(Document doc, DocumentBuilder doc_builder, 
            List<QestionReport> files) {
            int section_index = InsertionPos.csdgws_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table, 
                InsertionPos.csdgws_sec_table2, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(InsertionPos.csdgws_sec_table2, index_row, InsertionPos.csdgns_solu_row, 0);
                doc_builder.Write(report.chulicuoshi);
                index_row += 1;
            }
            return true;
        }

        //配置管理计划
        public void write_pzgljh_chart(Document doc, DocumentBuilder doc_builder) {
            List<FileList> files = ContentFlags.beicejianqingdan;
        }


        //common_func
        public bool write_file_list(Document doc, DocumentBuilder doc_builder, int sec_index,
            int sec_table_index, List<FileList> files, List<string> titles) {
            if(files.Count == 0){
                MessageBox.Show("请先上传项目基本信息文件");
            }
            int section_index = sec_index + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = 14;//???param
            foreach(FileList file in files) {
                if(index_row < files.Count + 13) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_name_row, 0);
                doc_builder.Write(file.wd_mingcheng);
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_iden_row, 0);
                doc_builder.Write(titles[index_row - 14]);
                index_row += 1;
            }
            return true;
        }

        public void write_pzgg_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files) {
            foreach(QestionReport report in files) {
                string temp = report.wenti;
                this.questions += temp + '\n';
            }
            if(doc_builder.MoveToBookmark("配置更改问题"))
                doc_builder.Write(this.questions);
        }
    }
}
