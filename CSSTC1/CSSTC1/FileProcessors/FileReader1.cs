using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors;
using CSSTC1.FileProcessors.models;

namespace CSSTC1.FileProcessors {
    public class FileReader1 {
        public static int count = 0;
        private FileWriters1 file_writer = new FileWriters1();
        public void read_charts(string filepath) {
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

            this.read_csry_chart(count, doc);
            count += 1;
            this.read_xmjs_chart(count, doc);
            count += 1;

            int[] index = { count, count + 1, count + 2 };
            this.read_wdqd_chart(index, doc);
            file_writer.conference_signing();
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
                foreach(Cell cell in row.Cells) {
                    string cellText = cell.GetText().ToString().Trim();
                    cellTexts.Add(cellText.Substring(0, cellText.Length - 1));
                }
                NamingRules.table_names.Add(index, "配置项测试");
                cellTexts.RemoveAt(0);
                file_writer.write_pzx_chart(cellTexts);
            }
        }



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

        public void read_xmjs_chart(int index, Document doc) {
            List<ProjectInfo> pro_infos = new List<ProjectInfo>();
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
                }

            }
            file_writer.write_xmjj_chart(pro_infos);
        }

        public void read_wdqd_chart(int[] index, Document doc) {
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
            file_writer.write_wdqd_chart(file_lists);

        }

    }
}

