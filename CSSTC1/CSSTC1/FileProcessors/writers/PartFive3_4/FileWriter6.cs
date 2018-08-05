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

namespace CSSTC1.FileProcessors.writers.PartFive3_4 {
    class FileWriter6 {
        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict;
        Dictionary<string, List<HardwareItems>> yingjianpeizhi_dict;
        List<SoftwareItems> softwares = new List<SoftwareItems>();
        int[] times = { ContentFlags.pianli_1, ContentFlags.pianli_2, ContentFlags.lingqucishu * 2 };
        public FileWriter6(Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string, 
            List<HardwareItems>> yingjianpeizhi_dict) {
            this.ruanjianpeizhi_dict = ruanjianpeizhi_dict;
            this.yingjianpeizhi_dict = yingjianpeizhi_dict;
            this.write_charts();
        }
        public void write_charts(){
            Document doc = new Document(FilePaths.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            this.write_bcjqd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.cssmps_bcjqd_section, InsertionPos.cssmps_bcjqd_sec_table,
                InsertionPos.cssmns_bcjqd_name_row, InsertionPos.cssmns_bcjqd_iden_row, times);
            doc.Save(FilePaths.save_root_file);
            MessageBox.Show("搭建环境就绪评审完成");
        }

        //测试环境调拨单表格
        public void write_bcjqd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string, List<HardwareItems>>
            yingjianpeizhi_dict, int section_index, int sec_table_index, int name_row_index, int iden_row_index, 
            int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;
            int merge_cell = 1;
            List<List<HardwareItems>> hardwares = yingjianpeizhi_dict.Select(r => r.Value).ToList();
            List<List<SoftwareItems>> softwares = ruanjianpeizhi_dict.Select(r => r.Value).ToList();
            for(int i = 0; i < yingjianpeizhi_dict.Count; i ++) {//循环每个软件的软件配置项和硬件配置项               
                foreach(HardwareItems key in hardwares[i]){ //某个软件的硬件配置
                //if(row_index < ruanjianpeizhi_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                    //}
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                    string name = key.yj_mingcheng;
                    doc_builder.Write(name);
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index + 1, 0);
                    doc_builder.Write("硬件设备");
                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = key.yj_bianhao;
                    doc_builder.Write(id);
                    doc_builder.MoveToCell(sec_table_index, 1,
                            InsertionPos.sj_bcjqd_date_row, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                    doc_builder.MoveToCell(sec_table_index, row_index,
                        InsertionPos.sj_bcjqd_date_row, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    Cell pre_cell = table.Rows[merge_cell].Cells[InsertionPos.sj_bcjqd_orig_row];
                    string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                    if(temp.Equals(key.wd_laiyuan)) {
                        //合并来源列
                        doc_builder.MoveToCell(sec_table_index, merge_cell,
                                InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                        doc_builder.MoveToCell(sec_table_index, row_index,
                            InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;

                    }
                    else {
                        doc_builder.MoveToCell(sec_table_index, row_index,
                            InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.Write(key.wd_laiyuan);
                        merge_cell = row_index;
                    }
                    row_index += 1;
                }
                Dictionary<string, SoftwareItems> files = this.set_file_id(doc, doc_builder, softwares[i]);
                int count = 0;
                List<string> file_ids = files.Select(r => r.Key).ToList();
                doc_builder.MoveToSection(cur_section);

                foreach(SoftwareItems key in softwares[i]) {  //某个软件的软件配置
                    //if(row_index < ruanjianpeizhi_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                    //}
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                    string name = key.rj_mingcheng;
                    doc_builder.Write(name);
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index + 1, 0);
                    doc_builder.Write("磁介质");
                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = file_ids[count];
                    doc_builder.Write(id);
                    doc_builder.MoveToCell(sec_table_index, 1,
                            InsertionPos.sj_bcjqd_date_row, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                    doc_builder.MoveToCell(sec_table_index, row_index,
                        InsertionPos.sj_bcjqd_date_row, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    Cell pre_cell = table.Rows[merge_cell].Cells[InsertionPos.sj_bcjqd_orig_row];
                    string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                    if(temp.Equals(key.wd_laiyuan)) {
                        //合并来源列
                        doc_builder.MoveToCell(sec_table_index, merge_cell,
                                InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                        doc_builder.MoveToCell(sec_table_index, row_index,
                            InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;

                    }
                    else {
                        doc_builder.MoveToCell(sec_table_index, row_index,
                            InsertionPos.sj_bcjqd_orig_row, 0);
                        doc_builder.Write(key.wd_laiyuan);
                        merge_cell = row_index;
                    }
                    row_index += 1;
                    count += 1;
                }
        }
            int sum = table.Rows.Count;
            table.Rows[sum - 1].Remove();
        }

        public Dictionary<string, SoftwareItems> set_file_id(Document doc, DocumentBuilder doc_builder, 
            List<SoftwareItems> software_items) {
            Dictionary<string, SoftwareItems> dict = new Dictionary<string, SoftwareItems>();
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
                key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C26";
                if(count < 10) {
                    key += "(0" + count.ToString() + ')' + "-" + item.rj_banben + '-' + year;
                }
                else
                    key += '(' + count.ToString() + ')' + "-" + item.rj_banben + '-' + year;
                dict.Add(key, item);
                count += 1;
            }
            return dict;
        }
    }

}
