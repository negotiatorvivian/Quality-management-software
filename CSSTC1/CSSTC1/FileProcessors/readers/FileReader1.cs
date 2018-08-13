using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors;
using CSSTC1.FileProcessors.models;
using CSSTC1.FileProcessors.writers;
using System.Windows.Forms;
using CSSTC1.CommonUtils;

namespace CSSTC1.FileProcessors.readers {
    public class FileReader1 {
        public static int count = 0;
        private FileWriter1 file_writer;

        public void read_charts(string filepath){
            if(ContentFlags.pingshenzuchengyuan.Count > 0){
                Document doc = new Document(filepath);
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                this.file_writer = new FileWriter1();
                if(ContentFlags.peizhiceshi > 0) {
                    this.read_pzx_chart(doc, doc_builder, (Table)tables[count], count);
                    count += 1;
                }
                else {
                    Table table = new Table(doc);
                    this.read_pzx_chart(doc, doc_builder, table, count);
                }
                if(ContentFlags.xitongceshi > 0) {
                    List<string> cellTexts = this.read_xt_chart((Table)tables[count], count);
                    count += 1;
                    file_writer.write_xt_chart(cellTexts);

                }
                else {
                    List<string> cellTexts = this.read_xt_chart((Table)tables[count], ContentFlags.missing);
                    file_writer.write_xt_chart(cellTexts);
                }
                file_writer.write_rwtzd_chart();

                string names = this.read_csry_chart((Table)tables[count], count);
                file_writer.write_csry_chart(names);

                count += 1;
                
                bool res = this.read_xmjs_chart((Table)tables[count]);
                count += 1;
                if(res){
                    for(int i = 0; i < ContentFlags.lingqucishu; i++){
                        int[] index = { count + i * 3, count + i * 3 + 1, count + i * 3 + 2 };
                        this.read_wdqd_chart(index, doc, i);
                    }
                }
                MessageBox.Show("写入文档完成!");
            }
            else
                MessageBox.Show("请检查上传表格内容！");

        }

        //配置项表格
        public void read_pzx_chart(Document doc, DocumentBuilder doc_builder, Table table, int index) {
            List<string> cellTexts = new List<string>();
            if(table == null) {
                file_writer.write_pzx_chart(cellTexts);
            }
            else {
                Row row = table.Rows[0];
                int jtfx_index = 0;
                int dmzc_index = 0;
                int dmsc_index = 0;
                foreach(Cell cell in row.Cells) {
                    string cellText = cell.GetText().ToString().Trim();
                    cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                }
                NamingRules.table_names.Add(index, "配置项测试");
                cellTexts.RemoveAt(0);
                file_writer.write_pzx_chart(cellTexts);
                if(ContentFlags.jingtaifenxi > 0){
                    jtfx_index = this.read_chart_info(doc, doc_builder, row, "静态分析\a");
                    if(jtfx_index > 0)
                        ContentFlags.static_list = this.read_jtfx_chart(doc, doc_builder, table, jtfx_index);
                }
                if(ContentFlags.daimashencha > 0) {
                    dmsc_index = this.read_chart_info(doc, doc_builder, row, "代码审查\a");
                    if(dmsc_index > 0)
                        ContentFlags.dmsc_software_list = this.read_jtfx_chart(doc, doc_builder, table, 
                            dmsc_index);
                }
                if(ContentFlags.daimazoucha > 0) {
                    dmzc_index = this.read_chart_info(doc, doc_builder, row, "代码走查\a");
                    if(dmzc_index > 0)
                       ContentFlags.dmzc_software_list = this.read_jtfx_chart(doc, doc_builder, table, dmzc_index);
                }
            }
        }

        public int read_chart_info(Document doc, DocumentBuilder doc_builder, Row row, string str){
            int cell_index = 0;
            foreach(Cell cell in row.Cells) {
                string cellText = cell.GetText().ToString().Trim();
                if(!cellText.Equals(str))
                    cell_index += 1;
                else
                    break;
            }
            if(cell_index == row.Cells.Count)
                return ContentFlags.missing;
            return cell_index;
        }

        //静态分析表格
        public List<string> read_jtfx_chart(Document doc, DocumentBuilder doc_builder, Table table, int index) {
            int row_index = 1;
            List<String> software_name = new List<string>();
            foreach(Row row in table.Rows){
                if(row_index == table.Rows.Count)
                    break;
                doc_builder.MoveToCell(0, row_index, index, 0);
                string temp = doc_builder.CurrentNode.Range.Text;
                if(temp.Equals("√")){
                    string name = table.Rows[row_index].Cells[0].Range.Text;
                    software_name.Add(name.Substring(0, name.Length - 1));
                }
                row_index += 1;
            }
            return software_name;
        }

        //系统测试表格
        public List<string> read_xt_chart(Table table, int index) {
            List<string> cellTexts = new List<string>();
            if(index > 0){
                Row row = table.Rows[0];
                foreach(Cell cell in row.Cells) {
                    string cellText = cell.GetText().ToString().Trim();
                    cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                }
                NamingRules.table_names.Add(index, "系统测试");
                cellTexts.RemoveAt(0);
            }
            return cellTexts;
        }

        //测试人员
        public string read_csry_chart(Table table, int index) {
            string names = "";
            foreach(Row row in table.Rows) {
                Cell cell = row.Cells[1];
                string cellText = cell.GetText().ToString().Trim();
                cellText = cellText.Substring(0, cellText.Length - 1);
                //cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                names += cellText + "、";
            }
            NamingRules.table_names.Add(index, "测试人员");
            //cellTexts.RemoveAt(0);
            names = names.Substring(3, names.Length - 4);
            return names;
        }

        //项目简介表格
        public bool read_xmjs_chart(Table table) {
            List<ProjectInfo> pro_infos = new List<ProjectInfo>();
            List<StaticAnalysisFile> static_files = new List<StaticAnalysisFile>();
            List<StaticAnalysisFile> code_walkthrough_files = new List<StaticAnalysisFile>();
            List<StaticAnalysisFile> code_review_files = new List<StaticAnalysisFile>();

            for(int i = 0; i < table.Rows.Count; i++) {
                string[] temp = new string[10];
                if(i == 0) {
                    for(int j = 1; j < table.Rows[i].Cells.Count - 1; j++) {
                        Cell cell = table.Rows[i].Cells[j];
                        string cellText = cell.GetText().ToString().Trim();
                        cellText = cellText.Substring(0, cellText.Length - 1);
                        NamingRules.xmjj_names.Add(cellText, NamingRules.Xmjj_paras[j - 1]);

                    }
                }
                else {
                    for(int j = 1; j < table.Rows[i].Cells.Count - 1; j++) {
                        Cell cell = table.Rows[i].Cells[j];
                        string cellText = cell.GetText().ToString();
                        cellText = cellText.Substring(0, cellText.Length - 1);
                        temp[j - 1] = cellText;
                    }
                    ProjectInfo pro_info = new ProjectInfo(temp[0], temp[1], temp[2], temp[3], temp[4],
                                        temp[5], temp[6], temp[7], temp[8], temp[9]);
                    pro_infos.Add(pro_info);
                    if(ContentFlags.static_list.Contains(temp[1])){
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "静态分析源代码", temp[2], 
                            temp[5], temp[8]);
                        static_files.Add(static_file);
                    }
                    if(ContentFlags.dmsc_software_list.Contains(temp[1])) {
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "代码审查源代码", temp[2],
                            temp[5], temp[8]);
                        code_review_files.Add(static_file);
                    }
                    if(ContentFlags.dmzc_software_list.Contains(temp[1])) {
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "代码走查源代码", temp[2],
                            temp[5], temp[8]);
                        code_walkthrough_files.Add(static_file);
                    }
                }

            }
            ContentFlags.static_files = static_files;
            ContentFlags.dmsc_software_info = code_review_files;
            ContentFlags.dmzc_software_info = code_walkthrough_files;
            ContentFlags.pro_infos = pro_infos;
            bool res = file_writer.write_xmjj_chart(pro_infos);
            return res;
        }


        //文档清单表格
        public void read_wdqd_chart(int[] index, Document doc, int time) {
            List<FileList> file_lists = new List<FileList>();
            List<FileList> all_file_lists = new List<FileList>();
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            foreach(int i in index) {
                Node node = doc.GetChild(NodeType.Table, i, true);
                Table table = (Table)node;
                for(int j = 1; j < table.Rows.Count - 1; j++) {
                    string[] temp = new string[10];
                    Row row = table.Rows[j];
                    for(int k = 1; k < row.Cells.Count; k++) {
                        Cell cell = row.Cells[k];
                        string cellText = cell.GetText().ToString().Trim();
                        cellText = cellText.Substring(0, cellText.Length - 1);
                        temp[k - 1] = cellText;
                    }
                    if(temp[4].Length == 0) {
                        temp[4] = file_lists[file_lists.Count - 1].wd_laiyuan;
                    }
                    if(temp[4].Equals(NamingRules.ignore_file))
                        continue;
                    if(temp[2].Equals("/") || temp[2].Equals("系泊后版本"))
                        temp[2] = "V1.0";
                    FileList all_file = new FileList(temp[0], temp[1], temp[2], temp[3], temp[4]);
                    all_file_lists.Add(all_file);
                    if(temp[2].Length > 4)
                        temp[2] = temp[2].Substring(0, 4);
                    FileList file = new FileList(temp[0], temp[1], temp[2], temp[3], temp[4]);
                    file_lists.Add(file);
                }
            }
            file_writer.write_wdqd_chart(file_lists, time);
            ContentFlags.all_file_lists = all_file_lists;
        }

    }
}

