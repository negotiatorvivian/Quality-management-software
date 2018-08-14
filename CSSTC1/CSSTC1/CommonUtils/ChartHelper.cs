using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;

namespace CSSTC1.CommonUtils {
    public class ChartHelper {
        //被测件清单表格
        public static void write_bcjqd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, FileList> new_dict, int section_index, int sec_table_index,
            int name_row_index, int iden_row_index, List<int> time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff){
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;
            int merge_cell = 1;
            foreach(string key in new_dict.Keys) {
                if(row_index < new_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                string name = new_dict[key].wd_mingcheng;
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(key);
                doc_builder.MoveToCell(sec_table_index, 1,
                        InsertionPos.sj_bcjqd_date_row, 0);
                doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                doc_builder.MoveToCell(sec_table_index, row_index,
                    InsertionPos.sj_bcjqd_date_row, 0);
                doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                Cell pre_cell = table.Rows[merge_cell].Cells[InsertionPos.sj_bcjqd_orig_row];
                string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                if(temp.Equals(new_dict[key].wd_laiyuan)) {
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
                    doc_builder.Write(new_dict[key].wd_laiyuan);
                    merge_cell = row_index;
                }
                row_index += 1;
            }
        }

        //配置项回归测试阶段的被测件调拨单和被测件领取清单
        public static void write_bcjdbd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, FileList> new_dict, int section_index, int sec_table_index, int row_index,
            int name_row_index, int iden_row_index, List<int> time_diff) {
            int flag = row_index;
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string key in new_dict.Keys) {
                if(row_index < new_dict.Count + flag - 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                string name = new_dict[key].wd_mingcheng;
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(key);
                row_index += 1;
            }
        }
        
        //静态分析、代码走查与代码审查阶段的入库申请单
        public static void write_rksqd_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, StaticAnalysisFile> new_dict, int section_index, int sec_table_index, int row_index,
            int name_row_index, int code_row_index, int iden_row_index, List<int> time_diff) {
            int flag = row_index;
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string key in new_dict.Keys) {
                if(row_index < new_dict.Count + flag - 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                string name = new_dict[key].rj_mingcheng + new_dict[key].jtfx_fanwei;
                doc_builder.Write(name);
                if(code_row_index > 0){
                    doc_builder.MoveToCell(sec_table_index, row_index, code_row_index, 0);
                    doc_builder.Write(new_dict[key].bc_yuyan);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(key);
                row_index += 1;
            }
        }

        //静态分析、代码走查与代码审查阶段被测件清单表格、被测件领取清单表格
        public static void write_bcjqd2_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, StaticAnalysisFile> new_dict, int section_index, int sec_table_index,
            int name_row_index, int iden_row_index, List<int> time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 1;
            int merge_cell = 1;
            foreach(string key in new_dict.Keys) {
                if(row_index < new_dict.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                string name = new_dict[key].rj_mingcheng + new_dict[key].jtfx_fanwei;
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(key);
                doc_builder.MoveToCell(sec_table_index, 1,
                        InsertionPos.sj_bcjqd_date_row, 0);
                doc_builder.CellFormat.VerticalMerge = CellMerge.First;
                doc_builder.MoveToCell(sec_table_index, row_index,
                    InsertionPos.sj_bcjqd_date_row, 0);
                doc_builder.CellFormat.VerticalMerge = CellMerge.Previous;
                Cell pre_cell = table.Rows[merge_cell].Cells[InsertionPos.sj_bcjqd_orig_row];
                string temp = pre_cell.Range.Text.Substring(0, pre_cell.Range.Text.Length - 1);
                if(temp.Equals(new_dict[key].yz_danwei)) {
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
                    doc_builder.Write(new_dict[key].yz_danwei);
                    merge_cell = row_index;
                }
                row_index += 1;
            }
        }

        //静态分析、代码走查与代码审查阶段被测件调拨单(入库申请单)表格
        public static void write_bcjdbd2_chart(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, StaticAnalysisFile> new_dict, int section_index, int sec_table_index, int row_index,
            int name_row_index, int iden_row_index, List<int> time_diff) {
            int flag = row_index;
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            foreach(string key in new_dict.Keys) {
                if(row_index < new_dict.Count + flag - 1) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                string name = new_dict[key].rj_mingcheng + new_dict[key].jtfx_fanwei;
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, iden_row_index, 0);
                doc_builder.Write(key);
                row_index += 1;
            }
        }

        //填写测试工具或设备核查单
        public static void append_content(Document doc, DocumentBuilder doc_builder, Table table, 
            int section_index, int sec_table_index, int row_index, List<int> time_diff) {
            int cur_section = section_index;
            foreach(int i in time_diff) {
                cur_section += i;
            }
            doc_builder.MoveToSection(cur_section);
            int cell_index = 1;
            Node node = doc.ImportNode(table, true);
            Table parent_table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            parent_table.Rows[row_index].Cells[cell_index].AppendChild(node);

        }
    }
}
