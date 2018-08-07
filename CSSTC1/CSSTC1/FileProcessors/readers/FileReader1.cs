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
        private FileWriter1 file_writer = new FileWriter1();

        public void read_charts(string filepath){
            if(ContentFlags.pingshenzuchengyuan.Count > 0){
                Document doc = new Document(filepath);
                NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
                if(ContentFlags.peizhiceshi) {
                    this.read_pzx_chart(count, doc);
                    count += 1;
                }
                else {
                    this.read_pzx_chart(ContentFlags.missing, doc);
                }
                if(ContentFlags.xitongceshi) {
                    this.read_xt_chart(count, doc);
                    count += 1;
                }
                else {
                    this.read_xt_chart(ContentFlags.missing, doc);
                }
                file_writer.write_rwtzd_chart();

                this.read_csry_chart(count, doc);
                count += 1;
                bool res = this.read_xmjs_chart(count, doc);
                count += 1;
                if(res){
                    for(int i = 0; i < ContentFlags.lingqucishu; i++){
                        int[] index = { count + i * 3, count + i * 3 + 1, count + i * 3 + 2 };
                        this.read_wdqd_chart(index, doc, i);
                    }
                }
                Document doc1 = new Document(FileConstants.save_root_file);
                DocumentBuilder doc_builder = new DocumentBuilder(doc1);
                bool res1 = file_writer.write_hyqdb_chart(doc1, doc_builder);
                if(res1)
                    MessageBox.Show("写入文档完成!");
            }
            else
                MessageBox.Show("请先填写测试需求分析与策划阶段的信息！");

        }

        public void read_pzx_chart(int index, Document doc) {
            List<string> cellTexts = new List<string>();
            if(index < 0) {
                file_writer.write_pzx_chart(cellTexts);
            }
            else {
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
                Node node = doc.GetChild(NodeType.Table, index, true);
                Table table = (Table)node;
                Row row = table.Rows[0];
                int jtfx_index = 0;
                int dmzc_index = 0;
                int dmsc_index = 0;
                bool flag = false;
                bool flag1 = false;
                bool flag2 = false;
                foreach(Cell cell in row.Cells) {
                    string cellText = cell.GetText().ToString().Trim();
                    cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                    if(!cellText.Equals("静态分析\a") && !flag)
                        jtfx_index += 1;
                    else if(cellText.Equals("静态分析\a"))
                        flag = true;
                    if(!cellText.Equals("代码审查\a") && !flag1)
                        dmsc_index += 1;
                    else if(cellText.Equals("代码审查\a"))
                        flag1 = true;
                    if(!cellText.Equals("代码走查\a") && !flag2)
                        dmzc_index += 1;
                    else if(cellText.Equals("代码走查\a"))
                        flag2 = true;
                }
                NamingRules.table_names.Add(index, "配置项测试");
                cellTexts.RemoveAt(0);
                file_writer.write_pzx_chart(cellTexts);
                this.read_jtfx_chart(doc, doc_builder, table, jtfx_index, ContentFlags.software_list);
                this.read_jtfx_chart(doc, doc_builder, table, dmsc_index, ContentFlags.dmsc_software_list);
                this.read_jtfx_chart(doc, doc_builder, table, dmzc_index, ContentFlags.dmzc_software_list);
            }
        }

        public void read_jtfx_chart(Document doc, DocumentBuilder doc_builder, Table table, int index,
            List<string> software_list) {
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
            software_list = software_name;
        }

        //public void read_dmsc_chart(Document doc, DocumentBuilder doc_builder, Table table, int index) {
        //    int row_index = 1;
        //    List<String> software_name = new List<string>();
        //    foreach(Row row in table.Rows) {
        //        if(row_index == table.Rows.Count)
        //            break;
        //        doc_builder.MoveToCell(0, row_index, index, 0);
        //        string temp = doc_builder.CurrentNode.Range.Text;
        //        if(temp.Equals("√")) {
        //            string name = table.Rows[row_index].Cells[0].Range.Text;
        //            software_name.Add(name.Substring(0, name.Length - 1));
        //        }
        //        row_index += 1;
        //    }
        //    ContentFlags.software_list = software_name;
        //}

        public void read_xt_chart(int index, Document doc) {
            List<string> cellTexts = new List<string>();
            if(index < 0) {
                file_writer.write_xt_chart(cellTexts);
            }
            else {
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
                Node node = doc.GetChild(NodeType.Table, index, true);
                Table table = (Table)node;
                Row row = table.Rows[0];
                foreach(Cell cell in row.Cells) {
                    string cellText = cell.GetText().ToString().Trim();
                    cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                }
                NamingRules.table_names.Add(index, "系统测试");
                cellTexts.RemoveAt(0);
                file_writer.write_xt_chart(cellTexts);
            }
        }

        public void read_csry_chart(int index, Document doc) {
            string names = "";
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Node node = doc.GetChild(NodeType.Table, index, true);
            Table table = (Table)node;
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
            file_writer.write_csry_chart(names);
        }

        public bool read_xmjs_chart(int index, Document doc) {
            List<ProjectInfo> pro_infos = new List<ProjectInfo>();
            List<StaticAnalysisFile> static_files = new List<StaticAnalysisFile>();
            List<StaticAnalysisFile> code_walkthrough_files = new List<StaticAnalysisFile>();
            List<StaticAnalysisFile> code_review_files = new List<StaticAnalysisFile>();
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Node node = doc.GetChild(NodeType.Table, index, true);
            Table table = (Table)node;

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
                    for(int j = i; j < table.Rows[i].Cells.Count - 1; j++) {
                        Cell cell = table.Rows[i].Cells[j];
                        string cellText = cell.GetText().ToString().Trim();
                        cellText = cellText.Substring(0, cellText.Length - 1);
                        temp[j - 1] = cellText;
                    }
                    ProjectInfo pro_info = new ProjectInfo(temp[0], temp[1], temp[2], temp[3], temp[4],
                                        temp[5], temp[6], temp[7], temp[8], temp[9]);
                    pro_infos.Add(pro_info);
                    if(ContentFlags.software_list.Contains(temp[1])){
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "静态分析源代码", temp[2], 
                            temp[5], temp[8]);
                        static_files.Add(static_file);
                    }
                    if(ContentFlags.dmsc_software_list.Contains(temp[1])) {
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "代码审查源代码", temp[2],
                            temp[5], temp[8]);
                        code_walkthrough_files.Add(static_file);
                    }
                    if(ContentFlags.dmzc_software_list.Contains(temp[1])) {
                        StaticAnalysisFile static_file = new StaticAnalysisFile(temp[1], "代码走查源代码", temp[2],
                            temp[5], temp[8]);
                        code_review_files.Add(static_file);
                    }
                }

            }
            ContentFlags.static_files = static_files;
            ContentFlags.dmsc_software_info = code_review_files;
            ContentFlags.dmzc_software_info = code_walkthrough_files;

            bool res = file_writer.write_xmjj_chart(pro_infos);
            return res;
        }


        //文档清单表格
        public void read_wdqd_chart(int[] index, Document doc, int time) {
            List<FileList> file_lists = new List<FileList>();
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
                    if(temp[2].Length > 4)
                        temp[2] = temp[2].Substring(0, 4);
                    FileList file = new FileList(temp[0], temp[1], temp[2], temp[3], temp[4]);
                    file_lists.Add(file);
                }
            }
            file_writer.write_wdqd_chart(file_lists, time);

        }

    }
}

