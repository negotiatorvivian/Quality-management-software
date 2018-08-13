using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;
using System.Text.RegularExpressions;
using CSSTC1.InputProcessors;
using CSSTC1.CommonUtils;
using System.Windows.Forms;

namespace CSSTC1.FileProcessors.writers {
    class FileWriter1 {
        private Document doc;
        private DocumentBuilder doc_builder;

        public FileWriter1(){
            this.doc = new Document(FileConstants.save_root_file);
            this.doc_builder = new DocumentBuilder(doc);
        }

        //配置项测试
        public void write_pzx_chart(List<string> cells) {
            this.doc_builder.MoveToCell(InsertionPos.pz_table, InsertionPos.pz_row, InsertionPos.pz_cell0, 0);
            this.doc_builder.Font.Name = "Wingdings 2";
            this.doc_builder.Font.Size = 12.0;
            if(cells.Count == 0) {
                this.doc_builder.Write(char.ConvertFromUtf32(163));
                OperationHelper.input_negative(doc, this.doc_builder, "测试环境", "配置项");
            }
            else {
                this.doc_builder.Write(char.ConvertFromUtf32(82));
                OperationHelper.input_confirm(doc, this.doc_builder, "测试环境", "配置项");

            }
            this.doc_builder.MoveToCell(InsertionPos.pz_table, InsertionPos.pz_row, InsertionPos.pz_cell1, 0);
            Node node = doc.GetChild(NodeType.Table, InsertionPos.pz_table, true);
            Table table = (Table)node;
            CellCollection cell_collection = table.Rows[InsertionPos.pz_row].Cells;
            foreach(Cell cell in cell_collection) {
                NodeCollection nodes = cell.GetChildNodes(NodeType.Paragraph, true);
                foreach(Paragraph para in nodes) {
                    string temp = para.Range.Text.Substring(0, para.Range.Text.Length - 1);
                    if(cells.Contains(temp)) {
                        ContentFlags.test_types.Add(temp);
                        OperationHelper.input_confirm(doc, this.doc_builder, temp, "配置");
                    }
                    else
                        OperationHelper.input_negative(doc, this.doc_builder, temp, "配置");
                }
            }
            //doc.Save(FileConstants.save_root_file);
        }

        //系统测试
        public void write_xt_chart(List<string> cells) {
            this.doc_builder.MoveToCell(InsertionPos.xt_table, InsertionPos.xt_row, InsertionPos.xt_cell0, 0);
            this.doc_builder.Font.Name = "Wingdings 2";
            this.doc_builder.Font.Size = 12.0;
            if(cells.Count == 0) {
                this.doc_builder.Write(char.ConvertFromUtf32(163));
                OperationHelper.input_negative(this.doc, this.doc_builder, "测试环境", "系统");
            }
            else {
                this.doc_builder.Write(char.ConvertFromUtf32(82));
                OperationHelper.input_confirm(doc, doc_builder, "测试环境", "系统");

            }
            this.doc_builder.MoveToCell(InsertionPos.xt_table, InsertionPos.xt_row, InsertionPos.xt_cell1, 0);
            Node node = doc.GetChild(NodeType.Table, InsertionPos.xt_table, true);
            Table table = (Table)node;
            CellCollection cell_collection = table.Rows[InsertionPos.xt_row].Cells;
            foreach(Cell cell in cell_collection) {
                NodeCollection nodes = cell.GetChildNodes(NodeType.Paragraph, true);
                foreach(Paragraph para in nodes) {
                    string temp = para.Range.Text.Substring(0, para.Range.Text.Length - 1);
                    if(cells.Contains(temp)) {
                        OperationHelper.input_confirm(this.doc, this.doc_builder, temp, "系统");
                        ContentFlags.test_types.Add(temp);
                    }
                    else
                        OperationHelper.input_negative(this.doc, this.doc_builder, temp, "系统");
                }
            }
            doc.Save(FileConstants.save_root_file);
        }

        //任务通知单
        public void write_rwtzd_chart(){;
            bool pzx_flag = true;
            bool xt_flag = true;
            if(ContentFlags.peizhiceshi == 0)
                pzx_flag = false;
            if(ContentFlags.xitongceshi == 0)
                xt_flag = false;
            string insert_str = "";
            if(pzx_flag && xt_flag){
                 insert_str = "配置项测试和系统测试";
            }
            else if(pzx_flag == true && xt_flag == false)
                insert_str = "配置项测试";
            else if(pzx_flag == false && xt_flag == true)
                insert_str = "系统测试";
            else
                insert_str = "";
            if(doc_builder.MoveToBookmark("测试级别"))
                this.doc_builder.Write(insert_str);
            string types_str = "";
            int test_num = 0;
            foreach(string t in ContentFlags.test_types){
                types_str += t + '、';
                test_num += 1;
            }
            if(types_str.Length > 0)
                types_str = types_str.Substring(0, types_str.Length - 1);
            if(this.doc_builder.MoveToBookmark("测试类型"))
                this.doc_builder.Write(types_str);
            if(this.doc_builder.MoveToBookmark("测试类型个数"))
                this.doc_builder.Write(test_num.ToString());
            //this.doc.Save(FileConstants.save_root_file);
        }

        //测试人员
        public void write_csry_chart(String names) {
            if(this.doc_builder.MoveToBookmark("测试组成员"))
                this.doc_builder.Write(names);
        }

        //项目简介（总）
        public bool write_xmjj_chart(List<ProjectInfo> pro_infos) {
            var languanges = new HashSet<string>();
            var environments = new HashSet<string>();
            string runtime_env = "";
            string code_env = "";
            float line_count = 0;
            foreach(ProjectInfo p in pro_infos) {
                runtime_env += (p.rj_mingcheng + "\n" + p.yx_huanjing + "\n");
                code_env += p.kf_huanjing;
                string[] items = p.bc_yuyan.Split('/');
                string[] numbers = p.dm_guimo.Split('/');
                foreach(string str in items)
                    languanges.Add(str);

                for(int i = 0; i < numbers.Length; i++) {
                    string str = numbers[i];
                    Regex reg = new Regex("[0-9]*[.]+[0-9]+");
                    Match match = reg.Match(str);
                    float lines = float.Parse(match.Groups[0].Value);
                    line_count += lines;
                }
            }
            ContentFlags.code_line = line_count;
            if(this.doc_builder.MoveToBookmark("代码行数"))
                this.doc_builder.Write(line_count.ToString());
            if(this.doc_builder.MoveToBookmark("运行环境"))
                this.doc_builder.Write(runtime_env);
            foreach(string language in languanges) {
                if(language.Equals("C++"))
                    OperationHelper.input_confirm(doc, this.doc_builder, "cplus", "");
                else if(language.Equals("C#"))
                    OperationHelper.input_confirm(doc, this.doc_builder, "csharp", "");
                else
                    OperationHelper.input_confirm(doc, this.doc_builder, language, "");
            }
            for(int i = 0; i < NamingRules.kfhj_params.Length; i++) {
                string str = NamingRules.kfhj_params[i];
                if(code_env.IndexOf(str, 0) >= 0) {
                    OperationHelper.input_confirm(doc, this.doc_builder, NamingRules.kfhj_bookmarks[i], "");
                }
            }
            this.write_hyqdb_chart(this.doc, doc_builder);
            //doc.Save(FileConstants.save_root_file);
            return true;

        }

        //填写文档清单命名
        public void write_wdqd_chart(List<FileList> files, int time) {
            Bookmark mark = doc.Range.Bookmarks["项目标识"];
            string xm_biaoshi = mark.Text;
            Bookmark mark1 = doc.Range.Bookmarks["项目开始时间"];
            string xm_nianfen = mark1.Text.Substring(0, 4);
            string title = NamingRules.pre_name + '{' + xm_biaoshi + '}';
            List<string> content_list = new List<string>();
            List<string> type_list = new List<string>();
            Dictionary<string, int> file_counter = ContentFlags.file_counter;
            Dictionary<string, int> file_type_counter = ContentFlags.file_type_counter;

            foreach(FileList file in files) {
                string type = MappingHelper.get_file_type(file.wd_mingcheng);
                if(file_type_counter.ContainsKey(type)) {
                    file_type_counter[type] += 1;
                }
                else
                    file_type_counter.Add(type, 1);
                type_list.Add(type);
            }
            for(int i = 0; i < files.Count; i++) {
                FileList file = files[i];
                int value = file_type_counter[type_list[i]];
                if(value > 1) {
                    if(file_counter.ContainsKey(type_list[i])) {
                        file_counter[type_list[i]] += 1;
                    }
                    else
                        file_counter.Add(type_list[i], 1);
                    string type_count = "(0" + file_counter[type_list[i]].ToString() + ')';
                    if(file_counter[type_list[i]] > 9)
                        type_count = "(" + file_counter[type_list[i]].ToString() + ')';
                    string text = title + "-C" + type_list[i] + type_count + '-' + 
                        file.wd_banben + '-' + xm_nianfen;
                    content_list.Add(text);
                }
                else {
                    string text = title + "-C" + type_list[i] + '-' + file.wd_banben + '-' + xm_nianfen;
                    content_list.Add(text);
                }
            }
            ContentFlags.file_counter = file_counter;
            ContentFlags.file_type_counter = file_type_counter;
            Dictionary<string, FileList> beicejianqingdan_dict = new Dictionary<string, FileList>();
            Dictionary<string, string> rukuwendang_dict = ContentFlags.rukuwendang_dict;
            int j = 0;
            foreach(FileList file in files) {
                if(!ContentFlags.beicejianqingdan_dict.ContainsKey(content_list[j]))
                    ContentFlags.beicejianqingdan_dict.Add(content_list[j], file);
                beicejianqingdan_dict.Add(content_list[j], file);
                if(!rukuwendang_dict.ContainsKey(file.wd_mingcheng)) {
                    rukuwendang_dict.Add(file.wd_mingcheng, content_list[j]);
                }
                j = j + 1;
            }

            ContentFlags.rukuwendang_dict = rukuwendang_dict;
            this.write_bcjdbd_chart(this.doc, this.doc_builder, files, content_list, time);
            this.write_bcjqd_chart(this.doc, this.doc_builder, beicejianqingdan_dict, time);
            this.write_bcjlqqd_chart(this.doc, this.doc_builder, files, content_list, time);
            this.write_lxwtf_chart(this.doc, this.doc_builder, files, time);
            this.write_rksqd_chart(this.doc, this.doc_builder, files, content_list, time);
            //if(ContentFlags.beicejianqingdan_dict.Count > 0)
            //    ContentFlags.beicejianqingdan_dict = beicejianqingdan_dict.Union
            //        (ContentFlags.beicejianqingdan_dict).ToDictionary(pair => pair.Key, pair => pair.Value);
            //else
            //    ContentFlags.beicejianqingdan_dict = beicejianqingdan_dict;
            doc.Save(FileConstants.save_root_file);
        }

        //被测件调拨单
        public void write_bcjdbd_chart(Document doc, DocumentBuilder rootdoc_builder, List<FileList> files, 
             List<string> content_list, int time) {
            Node node = doc.GetChild(NodeType.Table, InsertionPos.bcjdbd_table + 5 * time, true);
            Table table = (Table)node;

            int cur_section = 2 * time + 1;
            rootdoc_builder.MoveToSection(cur_section);
            rootdoc_builder.MoveToCell(InsertionPos.bcjdbd_sec_table, 1, InsertionPos.bcjdbd_pois_row, 0);
            rootdoc_builder.CellFormat.VerticalMerge = CellMerge.First;
            int row_index = 1;
            foreach(FileList file in files) {

                if(row_index < files.Count) {
                    if(row_index > 1) {
                        rootdoc_builder.MoveToCell(InsertionPos.bcjdbd_sec_table, 1, 
                            InsertionPos.bcjdbd_pois_row, 0);
                        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.First;
                        rootdoc_builder.MoveToCell(InsertionPos.bcjdbd_sec_table, row_index,
                            InsertionPos.bcjdbd_pois_row, 0);
                        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(1 + row_index, row);
                }

                rootdoc_builder.MoveToCell(InsertionPos.bcjdbd_sec_table, row_index, 
                    InsertionPos.bcjdbd_res_row, 0);
                Cell cell = table.Rows[row_index].Cells[1];
                rootdoc_builder.Write(content_list[row_index - 1]);
                rootdoc_builder.MoveToCell(InsertionPos.bcjdbd_sec_table, row_index, 
                    InsertionPos.bcjdbd_name_row, 0);
                rootdoc_builder.Write(file.wd_mingcheng);

                row_index += 1;
            }

        }

        //被测件清单
        public void write_bcjqd_chart(Document doc, DocumentBuilder rootdoc_builder, Dictionary<string, FileList> 
            beicejianqingdan_dict, int time) {
            List<int> time_diff = new List<int>();
            time_diff.Add(0);
            int cur_section = 2 * time + 1;
            int sec_table_index = 1;
            ChartHelper.write_bcjqd_chart(doc, rootdoc_builder, beicejianqingdan_dict, cur_section, 
                sec_table_index, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row, time_diff);
            //rootdoc_builder.MoveToSection(cur_section);
            //int row_index = 1;
            //int merge_cell = 1;
            
            //foreach(FileList file in files) {
            //    if(row_index < files.Count) {
            //        var row = table.Rows[row_index].Clone(true);
            //        table.Rows.Insert(1 + row_index, row);
            //    }
            //    rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, row_index, 
            //        InsertionPos.lx_bcjqd_name_row, 0);
            //    rootdoc_builder.Write(file.wd_mingcheng);
            //    rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, row_index,
            //        InsertionPos.lx_bcjqd_res_row, 0);
            //    rootdoc_builder.Write(content_list[row_index - 1]);
            //    rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, row_index,
            //        InsertionPos.lx_bcjqd_orig_row, 0);
            //    Cell pre_cell = table.Rows[merge_cell].Cells[InsertionPos.lx_bcjqd_orig_row];
            //    string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
            //    string date = DateHelper.cal_time(TimeStamp.lingqushijian[time], 0);
            //    if(temp.Equals(file.wd_laiyuan)) {
            //        //合并来源列
            //        rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, merge_cell,
            //                InsertionPos.lx_bcjqd_orig_row, 0);
            //        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.First;
            //        rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, row_index,
            //            InsertionPos.lx_bcjqd_orig_row, 0);
            //        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.Previous;

            //        //合并接收日期列
            //        rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, merge_cell,
            //            InsertionPos.lx_bcjqd_date_row, 0);
            //        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.First;
            //        rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, row_index,
            //            InsertionPos.lx_bcjqd_date_row, 0);
            //        rootdoc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
            //    }
            //    else {
            //        rootdoc_builder.Write(file.wd_laiyuan);
            //        merge_cell = row_index;
            //        rootdoc_builder.MoveToCell(InsertionPos.lx_bcjqd_sec_table, merge_cell,
            //            InsertionPos.lx_bcjqd_date_row, 0);
            //        rootdoc_builder.Write(date);
            //    }

            //    row_index += 1;
            //}
        }

        //被测件领取清单
        public void write_bcjlqqd_chart(Document doc, DocumentBuilder rootdoc_builder, 
            List<FileList> files, List<string> content_list, int time) {
            
            Node node = doc.GetChild(NodeType.Table, InsertionPos.bcjlqqd_table + 5 * time, true);
            Table table = (Table)node;

            int cur_section = 2 * time + 1;
            rootdoc_builder.MoveToSection(cur_section);
            int row_index = 2;

            foreach(FileList file in files) {
                if(row_index < files.Count + 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index, row);
                }
                rootdoc_builder.MoveToCell(InsertionPos.bcjlqqd_sec_table, row_index, 
                    InsertionPos.bcjlqqd_name_row, 0);
                rootdoc_builder.Write(file.wd_mingcheng);
                rootdoc_builder.MoveToCell(InsertionPos.bcjlqqd_sec_table, row_index, 
                    InsertionPos.bcjlqqd_res_row, 0);
                rootdoc_builder.Write(content_list[row_index - 2]);

                row_index += 1;
            }
        }

        //入库申请单
        public void write_rksqd_chart(Document doc, DocumentBuilder rootdoc_builder, List<FileList> files, 
            List<string> content_list, int time) {
            Node node = doc.GetChild(NodeType.Table, InsertionPos.rksqd_table + 5 * time, true);
            Table table = (Table)node;

            int cur_section = 2 * time + 1;
            rootdoc_builder.MoveToSection(cur_section);
            int row_index = 5;

            foreach(FileList file in files) {
                if(row_index < files.Count + 4) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                rootdoc_builder.MoveToCell(InsertionPos.rksqd_sec_table, row_index, 
                    InsertionPos.rksqd_name_row, 0);
                rootdoc_builder.Write(file.wd_mingcheng);
                rootdoc_builder.MoveToCell(InsertionPos.rksqd_sec_table, row_index, 
                    InsertionPos.rksqd_iden_row, 0);
                rootdoc_builder.Write(content_list[row_index - 5]);

                row_index += 1;

            }
            //doc.Save(FileConstants.save_root_file);
        }

        //联系委托方、配置报告单
        public void write_lxwtf_chart(Document doc, DocumentBuilder rootdoc_builder, List<FileList> files,
            int time) {
            string lxwtf_filenames = "";
            string pzztbgd_filesnames = "";
            foreach(FileList file in files) {
                lxwtf_filenames += file.wd_mingcheng + '、';
                pzztbgd_filesnames += file.wd_mingcheng + '\n';
            }
            lxwtf_filenames = lxwtf_filenames.Substring(0, lxwtf_filenames.Length - 1);
            //第八页填写的内容
            Bookmark bookmark = doc.Range.Bookmarks["被测件文档清单"];
            if(bookmark != null) {
                doc_builder.MoveToBookmark("被测件文档清单");
                doc_builder.Write(lxwtf_filenames);
            }
            //配置状态报告单中的配置状态
            string temp = "被测件清单" + (time + 1).ToString();
            Bookmark bookmark1 = doc.Range.Bookmarks[temp];
            if(bookmark1 != null) {
                doc_builder.MoveToBookmark(temp);
                doc_builder.Write(pzztbgd_filesnames);
            }
            //doc.Save(FileConstants.save_root_file);

        }

        //会议签到表
        public bool write_hyqdb_chart(Document doc, DocumentBuilder doc_builder) {
            int section_index = InsertionPos.cshtqd_section;
            int sec_table_index = InsertionPos.cshtqd_sec_table;
            bool res = OperationHelper.conference_signing(doc, doc_builder, section_index, sec_table_index);
            if(!res){
                return false;
            }
            else
                return true;
        }

    }
}
