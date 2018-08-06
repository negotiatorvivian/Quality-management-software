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
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict;
        List<List<string>> files_ids = new List<List<string>>();
        List<SoftwareItems> softwares = new List<SoftwareItems>();
        public string software_names = "";
        int[] times = { ContentFlags.pianli_1, ContentFlags.pianli_2, ContentFlags.lingqucishu * 2 };
        public FileWriter6(Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string,
            List<DynamicHardwareItems>> yingjianpeizhi_dict) {
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
            this.write_bcjdbd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.cssmps_bcjqd_section, InsertionPos.cssmps_bcjdbd_sec_table,
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, times);
            this.write_bcjlqqd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.cssmps_bcjqd_section, InsertionPos.cssmps_bcjlqqd_sec_table,
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, 
                InsertionPos.cssmps_bcjlqqd_type_row, times);
            this.write_pzztbg_chart(doc, doc_builder, "配置项被测件名称");
            this.write_csgjhsb_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict, 
                InsertionPos.djhj_cshjqr_section, InsertionPos.djhj_cshjqr_sec_table, times);
            this.append_content(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict, 
                InsertionPos.djhj_cshjhc_section, InsertionPos.djhj_cshjhc_sec_table + 1, 1, times);
            doc.Save(FilePaths.save_root_file);
            MessageBox.Show("搭建环境就绪评审完成");
        }

        //测试环境调拨单表格
        public void write_bcjdbd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string, 
            List<DynamicHardwareItems>>yingjianpeizhi_dict, int section_index, int sec_table_index, 
            int name_row_index, int iden_row_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;

            List<List<DynamicHardwareItems>> hardwares = yingjianpeizhi_dict.Select(r => r.Value).ToList();
            List<List<SoftwareItems>> softwares = ruanjianpeizhi_dict.Select(r => r.Value).ToList();
            for(int i = 0; i < yingjianpeizhi_dict.Count; i++) {//循环每个软件的软件配置项和硬件配置项               
                foreach(DynamicHardwareItems key in hardwares[i]) { //某个软件的硬件配置
                    //if(row_index < ruanjianpeizhi_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                    //}
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                    string name = key.yj_mingcheng;
                    doc_builder.Write(name);

                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = key.yj_bianhao;
                    doc_builder.Write(id);
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
                    
                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = file_ids[count];
                    doc_builder.Write(id);
                    
                    row_index += 1;
                    count += 1;
                }
            }
            int sum = table.Rows.Count;
            table.Rows[sum - 1].Remove();
        }

        //测试环境清单表格
        public void write_bcjqd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, 
            Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict, int section_index, 
            int sec_table_index,  int name_row_index, int iden_row_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;
            int merge_cell = 1;
            List<List<DynamicHardwareItems>> hardwares = yingjianpeizhi_dict.Select(r => r.Value).ToList();
            List<List<SoftwareItems>> softwares = ruanjianpeizhi_dict.Select(r => r.Value).ToList();
            for(int i = 0; i < yingjianpeizhi_dict.Count; i ++) {//循环每个软件的软件配置项和硬件配置项               
                foreach(DynamicHardwareItems key in hardwares[i]) { //某个软件的硬件配置
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
        
        //更新软件配置项标识版本号+0.1
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
                this.software_names += item.rj_mingcheng + '、';
            }
            return dict;
        }

        //测试环境领取清单表格
        public void write_bcjlqqd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string, 
            List<DynamicHardwareItems>> yingjianpeizhi_dict, int section_index, int sec_table_index,
            int name_row_index, int iden_row_index, int type_row_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 2;

            List<List<DynamicHardwareItems>> hardwares = yingjianpeizhi_dict.Select(r => r.Value).ToList();
            List<List<SoftwareItems>> softwares = ruanjianpeizhi_dict.Select(r => r.Value).ToList();
            for(int i = 0; i < yingjianpeizhi_dict.Count; i++) {//循环每个软件的软件配置项和硬件配置项               
                foreach(DynamicHardwareItems key in hardwares[i]) { //某个软件的硬件配置
                    //if(row_index < ruanjianpeizhi_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                    //}
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                    string name = key.yj_mingcheng;
                    doc_builder.Write(name);

                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = key.yj_bianhao;
                    doc_builder.Write(id);
                    doc_builder.MoveToCell(sec_table_index, row_index, type_row_index, 0);

                    doc_builder.Write("硬件");
                    row_index += 1;
                }
                Dictionary<string, SoftwareItems> files = this.set_file_id(doc, doc_builder, softwares[i]);
                int count = 0;
                List<string> file_ids = files.Select(r => r.Key).ToList();
                this.files_ids.Add(file_ids);
                doc_builder.MoveToSection(cur_section);

                foreach(SoftwareItems key in softwares[i]) {  //某个软件的软件配置
                    //if(row_index < ruanjianpeizhi_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                    //}
                    doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                    string name = key.rj_mingcheng;
                    doc_builder.Write(name);

                    doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                    string id = file_ids[count];
                    doc_builder.Write(id);
                    doc_builder.MoveToCell(sec_table_index, row_index, type_row_index, 0);

                    doc_builder.Write("磁介质");
                    row_index += 1;
                    count += 1;
                }
            }
            int sum = table.Rows.Count;
            table.Rows[sum - 1].Remove();
        }

        //搭建环境配置状态报告单
        public void write_pzztbg_chart(Document doc, DocumentBuilder doc_builder, string mark) {
            this.software_names = this.software_names.Substring(0, software_names.Length - 1);
            if(doc_builder.MoveToBookmark(mark))
                doc_builder.Write(this.software_names);
        }

        //软件动态测试配置项工具或设备确认表
        public void write_csgjhsb_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> new_dict, Dictionary<string, 
            List<DynamicHardwareItems>> new_dict1, int section_index, int sec_table_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            
            List<List<DynamicHardwareItems>> hardwares = new_dict1.Select(r => r.Value).ToList();
            List<List<SoftwareItems>> softwares = new_dict.Select(r => r.Value).ToList();
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;
            int temp = 0;
            while(temp < new_dict.Count - 1) {
                Section new_cshjqr_section = (Section)doc_builder.CurrentSection.Clone(true);
                doc_builder.MoveToSection(cur_section + 1);
                Section old_sibling_section = doc.Sections[cur_section + 1];
                Section new_sibling_section = (Section)doc_builder.CurrentSection.Clone(true);
                doc.InsertAfter(new_sibling_section, old_sibling_section);
                doc.InsertAfter(new_cshjqr_section, old_sibling_section);
                doc_builder.MoveToSection(cur_section);
                temp += 1;
            }
            doc.Save(FilePaths.save_root_file);
                
            for(int i = 0; i < new_dict.Count; i++) { //循环每个软件的软件配置项和硬件配置项   
                int count = 0;
                
                if(i > 0){
                    cur_section += 2;
                    doc_builder.MoveToSection(cur_section);
                    table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
                }
                row_index = 1;
                foreach(SoftwareItems software in softwares[i]) {
                    for(int j = 0; j < 4; j++) {
                        var row = table.Rows[row_index + j].Clone(true);
                        table.Rows.Insert(row_index + 4 + j, row);
                    }
                    doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_name_row, 0);
                    doc_builder.Write(software.rj_mingcheng);
                    doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_iden_row, 0);
                    doc_builder.Write(this.files_ids[i][count]);
                    row_index += 4;
                    count += 1;
                }
                foreach(DynamicHardwareItems hardware in hardwares[i]) {
                    for(int j = 0; j < 4; j++) {
                        var row = table.Rows[row_index + j].Clone(true);
                        table.Rows.Insert(row_index + 4 + j, row);
                    }
                    //DynamicHardwareItems hardware = hardwares[i][count];
                    doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_name_row, 0);
                    doc_builder.Write(hardware.yj_mingcheng);
                    doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.sj_csgjqr_iden_row, 0);
                    doc_builder.Write(hardware.yj_bianhao);
                    row_index += 4;
                }
            }
            int sum = table.Rows.Count;
            table.Rows[sum - 1].Remove();
            ContentFlags.beiceruanjianshuliang = new_dict.Count;
        }

        //软件动态测试配置项工具或设备核查单
        public void append_content(Document doc, DocumentBuilder doc_builder, 
            Dictionary<string, List<SoftwareItems>> new_dict, Dictionary<string, List<DynamicHardwareItems>> 
            new_dict1, int section_index, int sec_table_index, int row_index, int[] time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            List<List<SoftwareItems>> softwares = new_dict.Select(r => r.Value).ToList();
            doc_builder.MoveToSection(cur_section);
            int cell_index = 1;
            int count = 0;
            
            foreach(List<SoftwareItems> software_list in softwares){
                int j = 0;
                sec_table_index = 1;
                if(count > 0){
                    cur_section += 2;
                    doc_builder.MoveToSection(cur_section);
                }
                Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
                int merge_cell = 1;
                row_index = 1;
                foreach(SoftwareItems software in software_list){
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index +1, row);
                    doc_builder.MoveToCell(sec_table_index, row_index, cell_index, 0);
                    doc_builder.Write(software.rj_mingcheng);
                    doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 1, 0);
                    doc_builder.Write(software.rj_banben);
                    doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 2, 0);
                    doc_builder.Write(software.rj_yongtu);
                    doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 3, 0);
                    doc_builder.Write(this.files_ids[count][j]);
                    doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 4, 0);
                    doc_builder.Write(software.wd_laiyuan);
                    Cell pre_cell = table.Rows[merge_cell].Cells[5];
                    string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                    if(temp.Equals(software.wd_laiyuan)) {
                        //合并来源列
                        doc_builder.MoveToCell(sec_table_index, merge_cell, 5, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                        doc_builder.MoveToCell(sec_table_index, row_index, 5, 0);
                        doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    else {
                        doc_builder.MoveToCell(sec_table_index, row_index, 5, 0);
                        doc_builder.Write(software.wd_laiyuan);
                        merge_cell = row_index;
                    }
                    j += 1;
                    row_index += 1;
                }
                int sum = table.Rows.Count;
                table.Rows[sum - 1].Remove();
            
            
            sec_table_index += 1;
            Table hard_table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, 
                true);
            List<List<DynamicHardwareItems>> hardwares = new_dict1.Select(r => r.Value).ToList();
            merge_cell = 1;
            row_index = 1;
            //foreach(List<DynamicHardwareItems> hardware_list in hardwares) {
            List<DynamicHardwareItems> hardware_list = hardwares[count];
            foreach(DynamicHardwareItems hardware in hardware_list) {
                var row = hard_table.Rows[row_index].Clone(true);
                hard_table.Rows.Insert(row_index + 1, row);
                doc_builder.MoveToCell(sec_table_index, row_index, cell_index, 0);
                doc_builder.Write(hardware.yj_mingcheng + '\n' + hardware.yj_peizhi);
                doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 1, 0);
                doc_builder.Write(hardware.yj_yongtu);
                doc_builder.MoveToCell(sec_table_index, row_index, cell_index + 2, 0);
                doc_builder.Write(hardware.yj_bianhao);
               
                Cell pre_cell = hard_table.Rows[merge_cell].Cells[4];
                string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                if(temp.Equals(hardware.wd_laiyuan)) {
                    //合并来源列
                    doc_builder.MoveToCell(sec_table_index, merge_cell, 4, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                    doc_builder.MoveToCell(sec_table_index, row_index, 4, 0);
                    doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                }
                else {
                    doc_builder.MoveToCell(sec_table_index, row_index, 4, 0);
                    doc_builder.Write(hardware.wd_laiyuan);
                    merge_cell = row_index;
                }
                row_index += 1;
            }
            int sum1 = hard_table.Rows.Count;
            hard_table.Rows[sum1 - 1].Remove();
            count += 1;

        }
            doc.Save(FilePaths.save_root_file);

        //}
        }
    
    }
}
