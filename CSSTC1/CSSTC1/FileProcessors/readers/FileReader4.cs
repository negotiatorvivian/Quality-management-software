using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;
using System.Windows;
using CSSTC1.FileProcessors.writers;
using CSSTC1.FileProcessors.writers.part3_4;
using CSSTC1.FileProcessors.writers.PartFive3_4;

namespace CSSTC1.FileProcessors.readers {
    public class FileReader4 {
        private FileWriter5 writer1;
        private FileWriter4 writer;
        

        public bool read_charts(string filepath){
            Document doc = new Document(filepath);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            //静态分析软件项
            List<SoftwareItems> software_items = new List<SoftwareItems>();
            Table software_table = new Table(doc);
            //静态分析硬件项
            List<HardwareItems> hardware_items = new List<HardwareItems>();
            Table hardware_table = new Table(doc);
            //测试说明内部评审表
            List<QestionReport> reports = new List<QestionReport>();
            //就绪存在问题意见表
            List<QestionReport> jxwt_reports = new List<QestionReport>();
            int count = 0;
            if(ContentFlags.jingtaifenxi > 0){
            //静态分析软件项表格和硬件项表格
                software_table = this.read_rjjtcshj_chart(doc, doc_builder, count, software_items);
                hardware_table = this.read_yjjtcshj_chart(doc, doc_builder, count + 1, hardware_items);
                count += 2;
            }
            //读取内审意见表
            Table smns_table = (Table)doc.GetChild(NodeType.Table, count, true);
            reports = this.read_smns_chart(smns_table);
            count += 1;
            if(ContentFlags.pianli_3 > 0){
                Table jxwt_table = (Table)doc.GetChild(NodeType.Table, count, true);
                jxwt_reports = this.read_smns_chart(jxwt_table);
                count += 1;
            }
            this.writer = new FileWriter4(ContentFlags.static_files, software_table,
                hardware_table);
            this.writer1 = new FileWriter5(reports);
            return true;
        }
    
        //软件静态测试和动态测试环境表格
        public Table read_rjjtcshj_chart(Document doc, DocumentBuilder doc_builder, int index, 
            List<SoftwareItems> software_items) {
            int row_index = 1;
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            for(int i = row_index; i < table.Rows.Count; i++){
                Row row = table.Rows[i];
                string name = row.Cells[1].Range.Text;
                name = name.Substring(0, name.Length - 1);
                string version = row.Cells[2].Range.Text;
                version = version.Substring(0, version.Length - 1);
                string use = row.Cells[3].Range.Text;
                use = use.Substring(0, use.Length - 1);
                string provider = row.Cells[4].Range.Text;
                provider = provider.Substring(0, provider.Length - 1);
                SoftwareItems item = new SoftwareItems(name, version, use, provider);
                software_items.Add(item);
            }
            return table;
        }

        //软件静态测试环境硬件项表格
        public Table read_yjjtcshj_chart(Document doc, DocumentBuilder doc_builder, int index,
            List<HardwareItems> hardware_items) {
            int row_index = 1;
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            for(int i = row_index; i < table.Rows.Count; i++) {
                Row row = table.Rows[i];
                string name = row.Cells[1].Range.Text;
                name = name.Substring(0, name.Length - 1);
                string id = row.Cells[2].Range.Text;
                id = id.Substring(0, id.Length - 1);
                
                string use = row.Cells[3].Range.Text;
                use = use.Substring(0, use.Length - 1);
                string status = row.Cells[4].Range.Text;
                status = status.Substring(0, status.Length - 1);
                string provider = row.Cells[5].Range.Text;
                provider = provider.Substring(0, provider.Length - 1);
                HardwareItems item = new HardwareItems(name, id, use, status, provider);
                hardware_items.Add(item);
            }
            return table;
        }


        //测试说明内审意见表
        public List<QestionReport> read_smns_chart(Table t) {
            List<QestionReport> reports = new List<QestionReport>();
            for(int i = 1; i < t.Rows.Count; i++) {
                Row row = t.Rows[i];
                int column_num = row.Cells.Count;
                string[] contents = new string[5];
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
    }
}
