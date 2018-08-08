using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using CSSTC1.FileProcessors.writers.part3_4;
using CSSTC1.FileProcessors.writers.PartFive3_4;
using CSSTC1.ConstantVariables;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Windows;

namespace CSSTC1.FileProcessors.readers {
    public class FileReader5 {
        private FileWriter6 writer2;

        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict = new Dictionary<string,
                List<SoftwareItems>>();
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict = new Dictionary<string,
            List<DynamicHardwareItems>>();
        public void read_charts(string filepath){
            Document doc = new Document(filepath);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            for(int i = 0; i < ContentFlags.static_files.Count; i++) {
                List<DynamicHardwareItems> temp1 = new List<DynamicHardwareItems>();
                List<SoftwareItems> temp2 = new List<SoftwareItems>();
                //Table table1 = (Table)doc.GetChild(NodeType.Table, 2 * i + 3, true);
                //Table table2 = (Table)doc.GetChild(NodeType.Table, 2 * i + 4, true);
                this.read_rjjtcshj_chart(doc, doc_builder, 2 * i + 3, temp2);
                this.read_yjdtcshj_chart(doc, doc_builder, 2 * i + 4, temp1);
                if(temp1 == null || temp2 == null) {
                    MessageBox.Show("读取软件配置项动态测试环境软件项和硬件项表格失败");
                    //return false;
                }
                else {
                    ruanjianpeizhi_dict.Add(ContentFlags.dynamic_list[i], temp2);
                    yingjianpeizhi_dict.Add(ContentFlags.dynamic_list[i], temp1);
                }
            }
            this.writer2 = new FileWriter6(ruanjianpeizhi_dict, yingjianpeizhi_dict);

        }

        //配置项动态测试软件环境表格
        public Table read_rjjtcshj_chart(Document doc, DocumentBuilder doc_builder, int index,
            List<SoftwareItems> software_items) {
            int row_index = 1;
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            for(int i = row_index; i < table.Rows.Count; i++) {
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

        //软件动态测试环境硬件项表格
        public Table read_yjdtcshj_chart(Document doc, DocumentBuilder doc_builder, int index,
            List<DynamicHardwareItems> dy_hardware_items) {
            int row_index = 1;
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            for(int i = row_index; i < table.Rows.Count; i++) {
                Row row = table.Rows[i];
                string name = row.Cells[1].Range.Text;
                name = name.Substring(0, name.Length - 1);
                string id = row.Cells[2].Range.Text;
                id = id.Substring(0, id.Length - 1);
                string count = row.Cells[3].Range.Text;
                count = count.Substring(0, count.Length - 1);
                string use = row.Cells[4].Range.Text;
                use = use.Substring(0, use.Length - 1);
                string configuration = row.Cells[5].Range.Text;
                configuration = configuration.Substring(0, configuration.Length - 1);
                string status = row.Cells[6].Range.Text;
                status = status.Substring(0, status.Length - 1);
                string provider = row.Cells[7].Range.Text;
                provider = provider.Substring(0, provider.Length - 1);
                DynamicHardwareItems item = new DynamicHardwareItems(name, id, count, use, configuration,
                    status, provider);
                dy_hardware_items.Add(item);
            }
            return table;
        }
    }
}
