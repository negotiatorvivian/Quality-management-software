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


namespace CSSTC1.FileProcessors.writers.ConfigurationItemTest {
    class ConfigChartHelper {
        //软件测试环境标识命名
        public static Dictionary<string, FileList> set_file_id(Document doc, DocumentBuilder doc_builder,
            List<ProjectInfo> software_items) {
                Dictionary<string, FileList> dict = new Dictionary<string, FileList>();
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
            foreach(ProjectInfo item in software_items) {
                string key = NamingRules.pre_name;
                key += '{' + doc.Range.Bookmarks["项目标识"].Text + "}-C26";
                if(count < 10) {
                    key += "(0" + count.ToString() + ')' + "-" + item.xt_banben.Split('/')[0] + '-' + year;
                }
                else
                    key += '(' + count.ToString() + ')' + "-" + item.xt_banben.Split('/')[0] + '-' + year;
                count += 1;
                FileList file = new FileList(key, item.rj_mingcheng, item.xt_banben, "", 
                    item.yz_danwei);
                dict.Add(key, file);
                //this.software_names += item.rj_mingcheng + '、';
            }
            return dict;
        }


        public void write_rksqd_chart(Document doc, DocumentBuilder doc_builder, int section,
            int sec_table_index, int name_row_index, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
            }
            string software_name = doc.Range.Bookmarks["软件名称"].Text;
            List<string> contents = new List<string>();
            List<string> contents_new = new List<string>();
            int count = 1;
            string content = "";
            string content_new = "";
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态测试原始记录";
                count += 1;
                contents.Add(text);
                content += text + "V1.0" + '\n';
            }
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态回归测试用例集";
                count += 1;
                contents_new.Add(text);
                content_new += text + "V1.0" + '\n';
            }
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 4;
            foreach(string temp in contents) {
                if(row_index < contents.Count + 3) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            if(doc_builder.MoveToBookmark("配置项首轮动态测试配置状态报告单1"))
                doc_builder.Write(content);
        }

        public void write_rksqd_chart1(Document doc, DocumentBuilder doc_builder, int section,
            int sec_table_index, int name_row_index, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
            }
            string software_name = doc.Range.Bookmarks["软件名称"].Text;
            List<string> contents_new = new List<string>();
            int count = 1 + ContentFlags.pro_infos.Count;
            string content_new = "";
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "配置项动态回归测试用例集";
                count += 1;
                contents_new.Add(text);
                content_new += text + "V1.0" + '\n';
            }
            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 3;
            foreach(string temp in contents_new) {
                if(row_index < contents_new.Count + 2) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            if(doc_builder.MoveToBookmark("配置项回归测试配置状态报告单1"))
                doc_builder.Write(content_new);
        }

        public void write_rksqd_chart2(Document doc, DocumentBuilder doc_builder, int section,
            int sec_table_index, int name_row_index, List<int> time_diff) {
            foreach(int i in time_diff) {
                section += i;
            }
            string software_name = doc.Range.Bookmarks["软件名称"].Text;
            List<string> contents_new = new List<string>();
            int count = 1;
            string content_new = "";
            foreach(ProjectInfo info in ContentFlags.pro_infos) {
                string text = software_name;
                text += "测试执行记录附件" + count.ToString() + "-";
                text += info.rj_mingcheng + "软件配置项动态回归测试原始记录";
                count += 1;
                contents_new.Add(text);
                content_new += text + "V1.1" + '\n';
            }
            content_new += software_name + "问题报告单V1.1\n";
            if(ContentFlags.luojiceshi > 0)
                content_new += software_name + "逻辑测试执行记录及结果文件 V1.0\n";
            content_new = content_new.Substring(0, content_new.Length - 1);

            doc_builder.MoveToSection(section);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            int row_index = 4;
            foreach(string temp in contents_new) {
                if(row_index < contents_new.Count + 3) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(row_index + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, name_row_index, 0);
                doc_builder.Write(temp);
                row_index += 1;
            }
            if(doc_builder.MoveToBookmark("配置项回归测试配置状态报告单2"))
                doc_builder.Write(content_new);
        }
    }
}
