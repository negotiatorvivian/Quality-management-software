using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using Aspose.Words.Tables;
using System.Windows.Forms;
using CSSTC1.CommonUtils;

namespace CSSTC1.FileProcessors.writers {
    class FileWriter2 {
        public string questions = "";
        public void write_charts(List<QestionReport> ns_report, List<QestionReport> ws_report) {
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            bool res = this.write_csdgnswt_chart(doc, doc_builder, ns_report);
            if(res)
                this.write_csdgnswtzz_chart(doc, doc_builder, ns_report);
            //doc.Save(FileConstants.save_root_file);
            bool res1 = this.write_csdgwswt_chart(doc, doc_builder, ws_report);
            if(res1)
                this.write_csdgwswtzz_chart(doc, doc_builder, ws_report);
            this.write_pghd_file_list(doc, doc_builder, ContentFlags.beicejianqingdan_dict, 
                InsertionPos.pzgljh_section,InsertionPos.pzgljh_sec_table, InsertionPos.pzgljh_pghd_row_index);

            int start_row = InsertionPos.pzgljh_xqjx_row_index;
            List<string> file_ids = ContentFlags.beicejianqingdan_dict.Select(r => r.Key).ToList();
            List<string> file_names = ContentFlags.beicejianqingdan_dict.Select(r => r.Value.wd_mingcheng).
                ToList();
            if(ContentFlags.lingqucishu == 1) {
                this.write_xqjx_file_list(doc, doc_builder, file_ids, file_names, InsertionPos.pzgljh_section, 
                    InsertionPos.pzgljh_sec_table, start_row, ContentFlags.beicejianqingdan_dict.Count, 2);
            }
            else{  //领取多次
                int sum = 0;
                for(int i = 0; i < ContentFlags.lingqucishu; i++) {
                    int count = ContentFlags.beicejianshuliang[i];
                    List<string> ids = file_ids;
                    List<string> names = file_names;
                    if(i > 0){
                        ids.RemoveRange(0, sum);
                        names.RemoveRange(0, sum);
                    }
                    sum += count;
                    int remove_rows = 0;
                    if(i == ContentFlags.lingqucishu - 1)
                        remove_rows = 2 - i;
                    start_row = this.write_xqjx_file_list(doc, doc_builder, ids, names,
                        InsertionPos.pzgljh_section, InsertionPos.pzgljh_sec_table, start_row, count, remove_rows);
            }
            }
           
            this.write_pzgg_chart(doc, doc_builder, ws_report, "配置更改问题");
            doc.Save(FileConstants.save_root_file);

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

        //配置管理计划(配管活动部分)
        public int write_pghd_file_list(Document doc, DocumentBuilder doc_builder,Dictionary<string, FileList> 
            beicejianqingdan_dict, int sec_index, int sec_table_index, int row_index) {
            if(beicejianqingdan_dict.Count == 0) {
                MessageBox.Show("请先上传项目基本信息文件");
            }
            int section_index = sec_index + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = row_index;//???param
            foreach(string title in beicejianqingdan_dict.Keys) {
                var row = table.Rows[index_row].Clone(true);
                table.Rows.Insert(index_row + 1, row);
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_name_row, 0);
                doc_builder.Write(beicejianqingdan_dict[title].wd_mingcheng);
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_iden_row, 0);
                doc_builder.Write(title);
                index_row += 1;
            }
            table.Rows[index_row].Remove();
            return index_row;
        }

        //配置管理计划(需求基线部分)
        public int write_xqjx_file_list(Document doc, DocumentBuilder doc_builder,List<string> ids,
            List<string> names, int sec_index, int sec_table_index, int row_index, int file_count,
            int remove_rows){
            int section_index = sec_index + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = row_index;//???param
            for(int i = 0; i < file_count; i++) {
                var row = table.Rows[index_row].Clone(true);
                table.Rows.Insert(index_row + 1, row);
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_name_row + 1, 0);
                doc_builder.Write(names[i]);
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.pzgljh_iden_row + 1, 0);
                doc_builder.Write(ids[i]);
                index_row += 1;
            }
            table.Rows[index_row].Remove();
            if(remove_rows > 0){
                for(int j = 0; j < remove_rows; j++)
                    table.Rows[index_row].Remove();
            }
            return index_row;
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
            int section_index = InsertionPos.csdgqd_section + ContentFlags.lingqucishu * 2 + ContentFlags.pianli_1;
            int sec_table_index = InsertionPos.csdgqd_sec_table;
            OperationHelper.fill_pszcy_info(doc_builder, names);
            bool res = OperationHelper.conference_signing(doc, doc_builder, section_index, sec_table_index);
            if(!res)
                MessageBox.Show("请先填写测试需求分析与策划阶段的信息");
            else
                doc.Save(FileConstants.save_root_file);
        }
    }
}
