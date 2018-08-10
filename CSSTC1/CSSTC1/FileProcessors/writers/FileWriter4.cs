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
        List<StaticAnalysisFile> software_items = new List<StaticAnalysisFile>();
        Table software_table;
        Table hardware_table;
        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict = new Dictionary<string,
                List<SoftwareItems>>();
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict = new Dictionary<string,
            List<DynamicHardwareItems>>();
        List<int> times = new List<int>();

        //构造方法
        public FileWriter4(List<StaticAnalysisFile> software_items, Table software_table, Table hardware_table) {
            this.software_items = software_items;
            this.hardware_table = hardware_table;
            this.software_table = software_table;
            times.Add(ContentFlags.pianli_1);
            times.Add(ContentFlags.pianli_2);
            times.Add(ContentFlags.lingqucishu * 2);
            this.write_charts();
        }

        public void write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Dictionary<string, FileList> beicejianqingdan_dict = this.get_wdsc_files
                (ContentFlags.beicejianqingdan_dict);
            /***************************文档审查开始*********************************/
            if(ContentFlags.wendangshencha > 0){
                this.write_wdsc_time_line(doc, doc_builder);
                ChartHelper.write_bcjqd_chart(doc, doc_builder, beicejianqingdan_dict, 
                    InsertionPos.sj_bcjqd_section,
                    InsertionPos.sj_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, 
                    InsertionPos.sj_bcjqd_iden_row,
                    times);
                ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, 
                    InsertionPos.sj_bcjdbd_section,
                    InsertionPos.sj_bcjdbd_sec_table, 1, 
                    InsertionPos.sj_bcjdbd_name_row,InsertionPos.sj_bcjdbd_iden_row,
                    times);
                ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, 
                    InsertionPos.sj_bcjdbd_section,
                    InsertionPos.sj_bcjlqqd_sec_table, 2, InsertionPos.sj_bcjdbd_name_row, 
                    InsertionPos.sj_bcjdbd_iden_row, times);
                ChartHelper.write_bcjdbd_chart(doc, doc_builder, beicejianqingdan_dict, 
                    InsertionPos.sj_rksqd_section,
                    InsertionPos.sj_rksqd_sec_table, 5, InsertionPos.sj_rksqd_name_row,
                    InsertionPos.sj_rksqd_iden_row, times);
                this.write_pzztbg_chart(doc, doc_builder, beicejianqingdan_dict, "文档审查入库软件文档");
            }
            times.Add(ContentFlags.wendangshencha);
            /***************************文档审查结束*********************************/

            /***************************静态分析开始*********************************/
            if(ContentFlags.jingtaifenxi > 0) {
                this.write_jtfx_time_line(doc, doc_builder);
                this.write_lxwtf_chart(doc, doc_builder, software_items, "委托方提供代码");
                Dictionary<string, StaticAnalysisFile> software_dict = this.set_file_id(doc, doc_builder, 
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
                this.write_pzztbg2_chart(doc, doc_builder, software_dict, "被测件清单3", false);
            
                Table new_table = (Table)this.software_table.Clone(true);
                ChartHelper.append_content(doc, doc_builder, new_table, InsertionPos.sj_cshjhc_section,
                    InsertionPos.sj_cshjhc_sec_table, 1, times);
                Table new_table1 = (Table)this.hardware_table.Clone(true);
                ChartHelper.append_content(doc, doc_builder, new_table1, InsertionPos.sj_cshjhc_section,
                    InsertionPos.sj_cshjhc_sec_table, 2, times);
                Dictionary<string, StaticAnalysisFile> new_software_dict = this.update_softwareId(software_dict);
                ChartHelper.write_bcjqd2_chart(doc, doc_builder, new_software_dict, InsertionPos.sj_bcjqd_section3,
                    InsertionPos.sj_bcjqd_sec_table3, InsertionPos.sj_bcjqd_name_row,
                    InsertionPos.sj_bcjqd_iden_row, times);
                ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict, 
                    InsertionPos.sj_bcjdbd_section3,
                    InsertionPos.sj_bcjdbd_sec_table3, 1, InsertionPos.sj_bcjdbd_name_row,
                    InsertionPos.sj_bcjdbd_iden_row, times);
                ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict,
                    InsertionPos.sj_bcjlqqd_section3, InsertionPos.sj_bcjlqqd_sec_table3, 2,
                    InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, times);
                ChartHelper.write_bcjdbd2_chart(doc, doc_builder, new_software_dict, 
                    InsertionPos.sj_rksqd_section3,
                    InsertionPos.sj_rksqd_sec_table3, 3, InsertionPos.sj_rksqd_name_row,
                    InsertionPos.sj_rksqd_iden_row, times);
                this.write_pzztbg2_chart(doc, doc_builder, software_dict, "被测件清单4", true);
                ContentFlags.software_dict = new_software_dict;
            }
            else{
                Dictionary<string, StaticAnalysisFile> software_dict = this.set_file_id(doc, doc_builder,
                    software_items);
                Dictionary<string, StaticAnalysisFile> new_software_dict = this.update_softwareId(software_dict);
                ContentFlags.software_dict = new_software_dict;
            }
            times.Add(ContentFlags.jingtaifenxi);
            /***************************静态分析结束*********************************/
            doc.Save(FileConstants.save_root_file);
        }

        #region  公共方法
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
 
        #endregion

        #region  文档审查方法
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

        //获取文档审查部分的文档列表
        public Dictionary<string, FileList> get_wdsc_files(Dictionary<string, FileList> beicejianqingdan_dict) {
            List<string> file_ids = beicejianqingdan_dict.Select(r => r.Key).ToList();
            Dictionary<string, FileList> wdsc_files = new Dictionary<string, FileList>();
            int count = 0;
            foreach(FileList file in ContentFlags.all_file_lists) {
                if(file.wd_banben.Length > 4) {
                    string new_version = file.wd_banben.Substring(5, file.wd_banben.Length - 5);
                    file.wd_banben = new_version;
                    string id = file_ids[count];
                    Regex reg = new Regex("[v|V][0-9][.][0-9]+");
                    MatchCollection matches = reg.Matches(id, 0);

                    if(matches == null)
                        continue;
                    string temp = matches[matches.Count - 1].Value;
                    int start_index = matches[matches.Count - 1].Index;
                    int end_index = start_index + temp.Length;

                    id = id.Substring(0, start_index) + new_version + id.Substring(end_index, id.Length -
                        end_index);
                    //id = id.Replace(temp, new_version);
                    wdsc_files.Add(id, file);
                }
                count += 1;
            }
            return wdsc_files;
        }
        
        public void write_wdsc_time_line(Document doc, DocumentBuilder doc_builder){
            DateHelper.fill_time_blank(doc, doc_builder, "文档审查结果记录入库时间", TimeStamp.wdscqr_time, 1);
            //文档审查确认时间
            if(doc_builder.MoveToBookmark("文档审查确认时间"))
                doc_builder.Write(TimeStamp.wdscqr_format_time);
            //文档审查回归时间
            if(doc_builder.MoveToBookmark("文档审查回归时间"))
                doc_builder.Write(TimeStamp.wdschg_format_time);
        }
        #endregion

        #region 静态分析方法
        //联系委托方
        public void write_lxwtf_chart(Document doc, DocumentBuilder doc_builder,
            List<StaticAnalysisFile> software_items, string mark) {
            string text = "";
            foreach(StaticAnalysisFile item in software_items) {
                text += item.rj_mingcheng + '、';
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写文件标识
        public Dictionary<string, StaticAnalysisFile> set_file_id(Document doc, DocumentBuilder doc_builder, 
            List<StaticAnalysisFile> software_items) {
            Dictionary<string, StaticAnalysisFile> dict = new Dictionary<string, StaticAnalysisFile>();
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
            foreach(StaticAnalysisFile item in software_items) {
                string key = NamingRules.pre_name;
                key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C19";
                if(count < 10){
                    key += "(0" + count.ToString() + ')' + "-" + item.xt_banben + '-' + year;
                }
                else
                    key += '(' + count.ToString() + ')' + "-" + item.xt_banben + '-' + year;
                dict.Add(key, item);
                count += 1;
            }
            return dict;
        }

        //配置状态报告单
        public void write_pzztbg2_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, StaticAnalysisFile> new_dict, string mark, bool huigui) {
            string text = "";
            foreach(string key in new_dict.Keys) {
                StaticAnalysisFile file = new_dict[key];
                if(huigui)
                    text += file.rj_mingcheng + "\t\t" + file.hg_banben + '\n';
                else
                    text += file.rj_mingcheng + "\t\t" + file.xt_banben + '\n';
            }
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(text);
        }

        //填写测试工具或设备表格
        public void write_csgjhsb_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, SoftwareItems> new_dict, int section_index, int sec_table_index, 
            List<int> time_diff) {
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
  

        //软件配置项版本增加
        public Dictionary<string, StaticAnalysisFile> update_softwareId(Dictionary<string, StaticAnalysisFile>
            softwareItems_dict) {
            Dictionary<string, StaticAnalysisFile> new_dict = new Dictionary<string, StaticAnalysisFile>();

            foreach(string key in softwareItems_dict.Keys) {
                Regex reg = new Regex("[v|V][0-9][.][0-9]+");
                MatchCollection matches = reg.Matches(key, 0);

                if(matches == null)
                    continue;
                string temp = matches[matches.Count - 1].Value;
                int start_index = matches[matches.Count - 1].Index;
                int end_index = start_index + temp.Length;

                string new_key = key.Substring(0, start_index) + softwareItems_dict[key].hg_banben +
                    key.Substring(end_index, key.Length - end_index);
                new_dict.Add(new_key, softwareItems_dict[key]);
            }
            return new_dict;
        }

        public void write_jtfx_time_line(Document doc, DocumentBuilder doc_builder) {
            //联系研制方提供被测软件源代码时间
            DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第五次", TimeStamp.jtfxsc_time, 7);
            //静态分析审查时间
            if(doc_builder.MoveToBookmark("静态分析审查时间"))
                doc_builder.Write(TimeStamp.jtfxsc_format_time);
            //静态分析结果入库时间
            DateHelper.fill_time_blank(doc, doc_builder, "静态分析结果入库时间", TimeStamp.jtfxqr_time, 1);
            //联系委托方第六次:研制方来测评中心确认静态分析问题时间
            DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第六次", TimeStamp.jtfxqr_time, 3);
            if(doc_builder.MoveToBookmark("静态分析回归时间"))
                doc_builder.Write(TimeStamp.jtfxhg_format_time);
        }
        #endregion
    }
}

