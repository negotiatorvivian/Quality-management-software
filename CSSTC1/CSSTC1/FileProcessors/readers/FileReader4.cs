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
using CSSTC1.FileProcessors.writers.CodeReview_Walkthrough;
using CSSTC1.FileProcessors.writers.SystemTest;

namespace CSSTC1.FileProcessors.readers {
    public class FileReader4 {
        private FileWriter5 writer1;
        private FileWriter4 writer;
        private FileWriter7 writer2;
        

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
                ContentFlags.jxwt_reports = jxwt_reports;
                count += 1;
            }
            if(ContentFlags.xitongceshi > 0){
                List<SoftwareItems> system_softwares = this.read_xtrjhj_chart(doc, doc_builder, count);
                count += 1;
                List<DynamicHardwareItems> system_hardwares = this.read_xtyjhj_chart(doc, doc_builder, count);
                count += 1;
                ContentFlags.system_softwares = system_softwares;
                ContentFlags.system_hardwares = system_hardwares;
            }
            //软件测试用例
            this.read_csyl_chart(doc, doc_builder, count, ContentFlags.pro_infos.Count);
            count += ContentFlags.pro_infos.Count;
            //软件测试问题
            Dictionary<string, List<TestProblem>> problem_dict = this.read_cswt_chart(doc, doc_builder, count,
                ContentFlags.pro_infos.Count);
            ContentFlags.problems = problem_dict;
            count += ContentFlags.pro_infos.Count;
            if(ContentFlags.xitongceshi > 0){
                //系统测试用例
                this.read_csyl_chart(doc, doc_builder, count, 1);
                count += 1;
                //系统测试问题
                Dictionary<string, List<TestProblem>> system_problem_dict = this.read_cswt_chart(doc, doc_builder, 
                    count, 1);
                ContentFlags.system_problems = system_problem_dict;
                count += 1;
            }
            ContentFlags.sys_time_dict = this.read_system_time_chart(doc, doc_builder, count);
            this.writer = new FileWriter4(ContentFlags.static_files, software_table,
                hardware_table);
            this.writer1 = new FileWriter5(reports);
            if(ContentFlags.daimashencha > 0) {
                this.writer2 = new FileWriter7("代码审查", hardware_table);
            }
            if(ContentFlags.daimazoucha > 0){
                this.writer2 = new FileWriter7("代码走查", hardware_table);
            }

            MessageBox.Show("信息填写结束!");
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
                SoftwareItems item = new SoftwareItems(name, version, use, provider, "");
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

        //系统测试软件环境表格
        public List<SoftwareItems> read_xtrjhj_chart(Document doc, DocumentBuilder doc_builder, int index) {
            List<SoftwareItems> software_items = new List<SoftwareItems>();
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
                string id = row.Cells[4].Range.Text;
                id = id.Substring(0, id.Length - 1);
                string provider = row.Cells[5].Range.Text;
                provider = provider.Substring(0, provider.Length - 1);
                SoftwareItems item = new SoftwareItems(name, version, use, provider, id);
                software_items.Add(item);
            }
            return software_items;
        }

        //系统动态测试环境硬件项表格
        public List<DynamicHardwareItems> read_xtyjhj_chart(Document doc, DocumentBuilder doc_builder, 
            int index) {
            int row_index = 1;
            List<DynamicHardwareItems> hardware_items = new List<DynamicHardwareItems>();
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            for(int i = row_index; i < table.Rows.Count; i++) {
                Row row = table.Rows[i];
                string name = row.Cells[1].Range.Text;
                name = name.Substring(0, name.Length - 1);
                string id = row.Cells[2].Range.Text;
                id = id.Substring(0, id.Length - 1);
                string num = row.Cells[3].Range.Text;
                num = num.Substring(0, num.Length - 1);
                string use = row.Cells[4].Range.Text;
                use = use.Substring(0, use.Length - 1);
                string configure = row.Cells[5].Range.Text;
                configure = configure.Substring(0, configure.Length - 1);
                string status = row.Cells[6].Range.Text;
                status = status.Substring(0, status.Length - 1);
                string provider = row.Cells[7].Range.Text;
                provider = provider.Substring(0, provider.Length - 1);
                DynamicHardwareItems item = new DynamicHardwareItems(name, id, num, use, configure, status, 
                    provider);
                hardware_items.Add(item);
            }
            return hardware_items;
        }

        //测试用例表格
        public void read_csyl_chart(Document doc, DocumentBuilder doc_builder, int index, int number) {
            Dictionary<string, List<TestExample>> examples = new Dictionary<string,List<TestExample>>();
            for(int i = 0; i < number; i++){
                List<TestExample> example_list = new List<TestExample>(); 
                string key = "";
                if(number == 1){
                    Bookmark mark = doc.Range.Bookmarks["软件名称"];
                    if(mark != null)
                        key = mark.Text;
                    }
                else
                    key = ContentFlags.pro_infos[i].rj_mingcheng;
                Table table = (Table)doc.GetChild(NodeType.Table, index + i, true);
                for(int j = 2; j < table.Rows.Count - 2; j ++){
                    Row row = table.Rows[j];
                    string type = row.Cells[0].Range.Text;
                    type = type.Substring(0, type.Length - 1);
                    string sjyl_num = row.Cells[1].Range.Text;
                    sjyl_num = sjyl_num.Substring(0, sjyl_num.Length - 1);
                    string sjyl_percent = row.Cells[2].Range.Text;
                    sjyl_percent = sjyl_percent.Substring(0, sjyl_percent.Length - 1);
                    string zxyl_num = row.Cells[3].Range.Text;
                    zxyl_num = zxyl_num.Substring(0, zxyl_num.Length - 1);
                    string wtgyl_num = row.Cells[4].Range.Text;
                    wtgyl_num = wtgyl_num.Substring(0, wtgyl_num.Length - 1);
                    TestExample temp = new TestExample(type, sjyl_num, sjyl_percent, zxyl_num, wtgyl_num);
                    example_list.Add(temp);
                }
                examples.Add(key, example_list);
            }
            if(ContentFlags.example_dict.Count > 0)
                ContentFlags.example_dict.Concat(examples);
            else
                ContentFlags.example_dict = examples;
        }

        //测试问题表格
        public Dictionary<string, List<TestProblem>> read_cswt_chart(Document doc, DocumentBuilder doc_builder, 
            int index, int number) {
            Dictionary<string, List<TestProblem>> problems = new Dictionary<string, List<TestProblem>>();
            for(int i = 0; i < number; i++){
                List<TestProblem> problem_list = new List<TestProblem>(); 
                string key = ContentFlags.pro_infos[i].rj_mingcheng;
                Table table = (Table)doc.GetChild(NodeType.Table, index + i, true);
                for(int j = 2; j < table.Rows.Count - 2; j += 4){
                    string type = table.Rows[j].Cells[0].Range.Text;
                    type = type.Substring(0, type.Length - 1);
                    List<List<int>> rows = new List<List<int>>();
                    
                    for(int k = j; k < 4 + j; k++) {
                        Row row = table.Rows[k];
                        List<int> temp_list = new List<int>();
                        string sjwt_num = row.Cells[2].Range.Text;
                        sjwt_num = sjwt_num.Substring(0, sjwt_num.Length - 1);
                        temp_list.Add(int.Parse(sjwt_num));
                        string cxwt_num = row.Cells[3].Range.Text;
                        cxwt_num = cxwt_num.Substring(0, cxwt_num.Length - 1);
                        temp_list.Add(int.Parse(cxwt_num));
                        string wdwt_num = row.Cells[4].Range.Text;
                        wdwt_num = wdwt_num.Substring(0, wdwt_num.Length - 1);
                        temp_list.Add(int.Parse(wdwt_num));
                        string qtwt_num = row.Cells[5].Range.Text;
                        qtwt_num = qtwt_num.Substring(0, qtwt_num.Length - 1);
                        temp_list.Add(int.Parse(qtwt_num));
                        string glwt_num = row.Cells[6].Range.Text;
                        glwt_num = glwt_num.Substring(0, glwt_num.Length - 1);
                        temp_list.Add(int.Parse(glwt_num));
                        string ylwt_num = row.Cells[7].Range.Text;
                        ylwt_num = ylwt_num.Substring(0, ylwt_num.Length - 1);
                        temp_list.Add(int.Parse(ylwt_num));
                        rows.Add(temp_list);
                    }
                    TestProblem temp = new TestProblem(type, rows[0], rows[1], rows[2], rows[3]);
                    problem_list.Add(temp);
                }
                problems.Add(key, problem_list);
            }
            return problems;
        }

        //系统时间表
        public Dictionary<string, string> read_system_time_chart(Document doc, DocumentBuilder doc_builder, 
           int index){
            Table table = (Table)doc.GetChild(NodeType.Table, index, true);
            Dictionary<string, string> sys_time_dict = new Dictionary<string,string>();
            string xqfx_time = table.Rows[1].Cells[1].Range.Text;
            xqfx_time = xqfx_time.Substring(0, xqfx_time.Length - 1);
            sys_time_dict.Add("需求分析阶段", xqfx_time);
            string csch_time = table.Rows[2].Cells[1].Range.Text;
            csch_time = csch_time.Substring(0, csch_time.Length - 1);
            sys_time_dict.Add("策划阶段", csch_time);
            string cssj_time = table.Rows[3].Cells[1].Range.Text;
            cssj_time = cssj_time.Substring(0, cssj_time.Length - 1);
            sys_time_dict.Add("设计阶段", cssj_time);
            int row_index1 = 4;
            for(int i = row_index1; i < table.Rows.Count - row_index1; i++) {
                Row row = table.Rows[i];
                Cell cur_cell = row.Cells[2];
                if(cur_cell.CellFormat.VerticalMerge == CellMerge.Previous){
                    row_index1 = i - 1;
                    break;
                }
            }
            for(int i = row_index1; i < 5 + row_index1; i++) {
                Row row = table.Rows[i];
                Cell pre_cell = row.Cells[1];
                Cell cur_cell = row.Cells[3];
                string value = pre_cell.GetText();
                string key = cur_cell.GetText();
                if(!sys_time_dict.ContainsKey(key))
                    sys_time_dict.Add(key, value);
                else 
                    break;
            }
            int count = table.Rows.Count - 3;
            if(ContentFlags.xitongceshi > 0) {
                string xtcs_time = table.Rows[count].Cells[1].GetText();
                sys_time_dict.Add("系统测试", xtcs_time);
            }
            string zjjd_time = table.Rows[count + 1].Cells[1].GetText();
            sys_time_dict.Add("总结阶段", zjjd_time);
            return sys_time_dict;
        }
    }
}
