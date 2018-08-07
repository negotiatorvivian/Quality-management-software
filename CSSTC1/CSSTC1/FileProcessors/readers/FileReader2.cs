using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using CSSTC1.FileProcessors.writers;
using CSSTC1.FileProcessors.models;
using System.Windows.Forms;
using CSSTC1.ConstantVariables;


namespace CSSTC1.FileProcessors.readers {
    public class FileReader2 {
        private FileWriter2 writer = new FileWriter2();
        public void read_charts(string filepath){
            Document doc = new Document(filepath);
            List<QestionReport> ns_reports = new List<QestionReport>();
            List<QestionReport> ws_reports = new List<QestionReport>();

            NodeCollection nodes = doc.GetChildNodes(NodeType.Table, true);
            if(nodes.Count != 2)
                MessageBox.Show("请检查上传文件内容");
            else{
                Table t0 = (Table)nodes[0];
                ns_reports = this.read_dg_chart(t0);
                Table t1 = (Table)nodes[1];
                ws_reports = this.read_dg_chart(t1);
                writer.write_charts(ns_reports, ws_reports);
            }
        }

        public List<QestionReport> read_dg_chart(Table t) {
            List<QestionReport> reports = new List<QestionReport>();
            for(int i = 1; i < t.Rows.Count; i ++){
                Row row = t.Rows[i];
                int column_num = row.Cells.Count;
                string[] contents= new string[5];
                for(int j = 1; j < column_num; j++) {
                    Cell cell = row.Cells[j];
                    string temp = cell.GetText().ToString();
                    contents[j - 1] = temp.Substring(0, temp.Length - 1);
                }
                QestionReport qus_report = new QestionReport(contents[0], contents[1], contents[2], contents[3]);
                reports.Add(qus_report);
            }
            return reports;
        }


        public void read_psz_members(List<string> names) {
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            writer.write_hyqdb_chart(doc, doc_builder, names);
        }

    }
}
