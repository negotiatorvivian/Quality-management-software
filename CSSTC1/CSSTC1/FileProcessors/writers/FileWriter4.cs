using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using System.Text.RegularExpressions;
using Aspose.Words;
using CSSTC1.CommonUtils;
using System.Windows;
using CSSTC1.FileProcessors.writers.part3_4;

namespace CSSTC1.FileProcessors.writers {
    class FileWriter4 {
        List<SoftwareItems> software_items = new List<SoftwareItems>();
        List<HardwareItems> hardware_items = new List<HardwareItems>();
        List<QestionReport> reports = new List<QestionReport>();
        Table software_table;
        Table hardware_table;
        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict = new Dictionary<string,
                List<SoftwareItems>>();
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict = new Dictionary<string,
            List<DynamicHardwareItems>>();

        public FileWriter4(List<SoftwareItems> software_items, List<HardwareItems> hardware_items, 
            List<QestionReport> reports, Table software_table, Table hardware_table, Dictionary<string, 
            List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string,
            List<DynamicHardwareItems>> yingjianpeizhi_dict) {
            this.hardware_items = hardware_items;
            this.software_items = software_items;
            this.hardware_table = hardware_table;
            this.software_table = software_table;
            this.yingjianpeizhi_dict = yingjianpeizhi_dict;
            this.ruanjianpeizhi_dict = ruanjianpeizhi_dict;
            this.reports = reports;
            this.write_charts();
        }

        public void write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Dictionary<string, FileList> beicejianqingdan_dict = this.update_fileId
                (ContentFlags.beicejianqingdan_dict);
            int[] times = { ContentFlags.pianli_1, ContentFlags.pianli_2, ContentFlags.lingqucishu *2};
            ChartHelper.write_bcjqd_chart(doc, doc_builder, beicejianqingdan_dict, InsertionPos.sj_bcjqd_section,
                InsertionPos.sj_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row,
                times);
            ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, InsertionPos.sj_bcjdbd_section,
                InsertionPos.sj_bcjdbd_sec_table, 1, 
                InsertionPos.sj_bcjdbd_name_row,InsertionPos.sj_bcjdbd_iden_row,
                times);
            ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, InsertionPos.sj_bcjdbd_section,
                InsertionPos.sj_bcjlqqd_sec_table, 2, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, times);
            ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, InsertionPos.sj_rksqd_section,
                InsertionPos.sj_rksqd_sec_table, 5, InsertionPos.sj_rksqd_name_row,
                InsertionPos.sj_rksqd_iden_row, times);
            this.write_pzztbg_chart(doc, doc_builder, beicejianqingdan_dict, "文档审查入库软件文档");
            this.write_lxwtf_chart(doc, doc_builder, software_items, "委托方提供代码");
            Dictionary<string, SoftwareItems> software_dict = this.set_file_id(doc, doc_builder, 
                software_items);
            ChartHelper.write_rksqd_chart(doc, doc_builder, software_dict, InsertionPos.sj_bcjdbd_section2,
                InsertionPos.sj_bcjdbd_sec_table2, 1, InsertionPos.sj_bcjdbd_name_row,
                InsertionPos.sj_bcjdbd_iden_row, times);
            ChartHelper.write_bcjqd2_chart(doc, doc_builder, software_dict, InsertionPos.sj_bcjqd_section2,
                InsertionPos.sj_bcjqd_sec_table2, InsertionPos.sj_bcjqd_name_row,
                InsertionPos.sj_bcjqd_iden_row, times);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, software_dict,
                InsertionPos.sj_bcjlqqd_section2, InsertionPos.sj_bcjlqqd_sec_table2, 2, 
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, times);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, software_dict, InsertionPos.sj_rksqd_section2,
                InsertionPos.sj_rksqd_sec_table2, 3, InsertionPos.sj_rksqd_name_row,
                InsertionPos.sj_rksqd_iden_row, times);
            this.write_pzztbg2_chart(doc, doc_builder, software_dict, "被测件清单3");
            this.write_csgjhsb_chart(doc, doc_builder, software_dict, InsertionPos.sj_csgjqr_section,
               InsertionPos.sj_csgjqr_sec_table, times);
            
            Table new_table = (Table)this.software_table.Clone(true);
            this.append_content(doc, doc_builder, new_table, InsertionPos.sj_cshjhc_section,
                InsertionPos.sj_cshjhc_sec_table, 1, times);
            Table new_table1 = (Table)this.hardware_table.Clone(true);
            this.append_content(doc, doc_builder, new_table1, InsertionPos.sj_cshjhc_section,
                InsertionPos.sj_cshjhc_sec_table, 2, times);
            Dictionary<string, SoftwareItems> new_software_dict = this.update_softwareId(software_dict);
            ChartHelper.write_bcjqd2_chart(doc, doc_builder, new_software_dict, InsertionPos.sj_bcjqd_section3,
                InsertionPos.sj_bcjqd_sec_table3, InsertionPos.sj_bcjqd_name_row,
                InsertionPos.sj_bcjqd_iden_row, times);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict, InsertionPos.sj_bcjdbd_section3,
                InsertionPos.sj_bcjdbd_sec_table3, 1, InsertionPos.sj_bcjdbd_name_row,
                InsertionPos.sj_bcjdbd_iden_row, times);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict,
                InsertionPos.sj_bcjlqqd_section3, InsertionPos.sj_bcjlqqd_sec_table3, 2,
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, times);
            ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict, InsertionPos.sj_rksqd_section3,
                InsertionPos.sj_rksqd_sec_table3, 3, InsertionPos.sj_rksqd_name_row,
                InsertionPos.sj_rksqd_iden_row, times);
            this.write_pzztbg2_chart(doc, doc_builder, software_dict, "被测件清单4");
            doc.Save(FileConstants.save_root_file);
            ContentFlags.software_dict = new_software_dict;
            MessageBox.Show("文档审查与静态分析部分写入完成！");
            //FileWriter5 append_writer = new FileWriter5(new_software_dict);
        }

        public Dictionary<string, FileList> update_fileId(Dictionary<string, FileList> beicejianqingdan_dict) {
            Dictionary<string, FileList> new_dict = new Dictionary<string,FileList>();
            foreach(string key in beicejianqingdan_dict.Keys){
                Regex reg = new Regex("[v|V][0-9][.][0-9]+");
                MatchCollection matches = reg.Matches(key, 0);
                
                if(matches == null)
                    continue;
                string temp = matches[matches.Count - 1].Value;
                string temp1 = temp.Substring(1, temp.Length - 1);
                double t = Double.Parse(temp1);
                t = t + 0.1;
                beicejianqingdan_dict[key].wd_banben = "V" + t.ToString();
                string id = key.Replace(temp, "V" + t.ToString());
                new_dict.Add(id, beicejianqingdan_dict[key]);
            }
            return new_dict;
        }

        public void write_pzztbg_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, FileList> new_dict, string mark){
            string text = "";
            foreach(string key in new_dict.Keys){
                FileList file = new_dict[key];
                text += file.wd_mingcheng + "\t\t" + file.wd_banben + '\n';
            }
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }
        //联系委托方
        public void write_lxwtf_chart(Document doc, DocumentBuilder doc_builder,
            List<SoftwareItems> software_items, string mark) {
            string text = "";
            foreach(SoftwareItems item in software_items) {
                text += item.rj_mingcheng + '、';
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写文件标识
        public Dictionary<string, SoftwareItems> set_file_id(Document doc, DocumentBuilder doc_builder, 
            List<SoftwareItems> software_items) {
            Dictionary<string, SoftwareItems> dict = new Dictionary<string,SoftwareItems>();
            string id = "";
            string year = "";
            if(doc_builder.MoveToBookmark("项目标识"))
                id = doc.Range.Bookmarks["项目标识"].Text;
            else
                return null;
            if(doc_builder.MoveToBookmark("年份"))
                year = doc.Range.Bookmarks["年份"].Text;
            else
                return null;
            int count = 1;
            foreach(SoftwareItems item in software_items) {
                string key = NamingRules.pre_name;
                key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C19";
                if(count < 10){
                    key += "(0" + count.ToString() + ')' + "-" + item.rj_banben + '-' + year;
                }
                else
                    key += '(' + count.ToString() + ')' + "-" + item.rj_banben + '-' + year;
                dict.Add(key, item);
                count += 1;
            }
            return dict;
        }

        //配置状态报告单
        public void write_pzztbg2_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, SoftwareItems> new_dict, string mark) {
            string text = "";
            foreach(string key in new_dict.Keys) {
                SoftwareItems file = new_dict[key];
                text += file.rj_mingcheng + "\t\t" + file.rj_banben + '\n';
            }
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写测试工具或设备表格
        public void write_csgjhsb_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, SoftwareItems> new_dict, int section_index, int sec_table_index, 
            int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 9;
            foreach(string key in new_dict.Keys){
                SoftwareItems software = new_dict[key];
                if(row_index / 4 < new_dict.Count + 1) {
                    for(int i = 0; i < 4; i ++){
                        var row = table.Rows[row_index + i].Clone(true);
                        table.Rows.Insert(row_index + 4 + i, row);
                    }
                    //doc.Save(FilePaths.save_root_file);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_name_row, 0);
                doc_builder.Write(software.rj_mingcheng);
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_iden_row, 0);
                doc_builder.Write(key);
                row_index += 4;
            }
        }
    

        //填写测试工具或设备核查单
        public void append_content(Document doc, DocumentBuilder doc_builder, Table table, int section_index, 
            int sec_table_index, int row_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            int cell_index = 1;
            //Section prepend_section = doc.Sections[cur_section];
            Node node = doc.ImportNode(table, true);
            Table parent_table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            parent_table.Rows[row_index].Cells[cell_index].AppendChild(node);

        }

        //软件配置项版本增加
        public Dictionary<string, SoftwareItems> update_softwareId(Dictionary<string, SoftwareItems>
            softwareItems_dict) {
                Dictionary<string, SoftwareItems> new_dict = new Dictionary<string, SoftwareItems>();
                foreach(string key in softwareItems_dict.Keys) {
                Regex reg = new Regex("[v|V][0-9][.][0-9]+");
                MatchCollection matches = reg.Matches(key, 0);

                if(matches == null)
                    continue;
                string temp = matches[matches.Count - 1].Value;
                string temp1 = temp.Substring(1, temp.Length - 1);
                double t = Double.Parse(temp1);
                t = t + 0.1;
                softwareItems_dict[key].rj_banben = "V" + t.ToString();
                string id = key.Replace(temp, "V" + t.ToString());
                new_dict.Add(id, softwareItems_dict[key]);
            }
            return new_dict;
        }
    }
}
