using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;
using System.Reflection;

namespace CSSTC1.CommonUtils {
    class OperationHelper {
        public static int copy = 0;
        public static void input_confirm(Document doc, DocumentBuilder doc_builder, string temp, string type) {
            Bookmark bookmark = doc.Range.Bookmarks[type + temp];
            if(bookmark != null) {
                bookmark.Text = "";
                doc_builder.MoveToBookmark(type + temp);
                doc_builder.Font.Name = "Wingdings 2";
                doc_builder.Font.Size = 12.0;
                doc_builder.Write(char.ConvertFromUtf32(82));
            }
        }

        public static void input_negative(Document doc, DocumentBuilder doc_builder, string temp, string type) {
            Bookmark bookmark = doc.Range.Bookmarks[type + temp];
            if(bookmark != null) {
                bookmark.Text = "";
                doc_builder.MoveToBookmark(type + temp, true, false);
                doc_builder.Font.Name = "Wingdings 2";
                doc_builder.Font.Size = 12.0;
                doc_builder.Write(char.ConvertFromUtf32(163));
            }

        }

        public static void delete_section(Document doc, DocumentBuilder doc_builder, string[] bookmarks) {
            foreach(string mark in bookmarks){
                if(doc_builder.MoveToBookmark(mark))
                    doc_builder.CurrentSection.Range.Delete();
            }
        }

        public static void delete_section(string path, string[] bookmarks) {
            Document doc = new Document(path);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            foreach(string mark in bookmarks) {
                if(doc_builder.MoveToBookmark(mark))
                    doc_builder.CurrentSection.Range.Delete();
            }
            doc.Save(path);
        }

        public static void delete_section(Document doc, DocumentBuilder doc_builder, string bookmark,
            string section_flag){
            BindingFlags flag = BindingFlags.Static | BindingFlags.Public;
            FieldInfo f_key = typeof(FileConstants).GetField(section_flag, flag);
            if(f_key != null){
                string temp = f_key.GetValue(new FileConstants()).ToString();
                int section_count = Int16.Parse(temp);
                List<Section> sections = new List<Section>();
                if(doc_builder.MoveToBookmark(bookmark)){
                    Section section = doc_builder.CurrentSection;
                    sections.Add(section);
                    for(int i = 0; i < section_count - 1; i++){
                        Section new_section = (Section)section.NextSibling;
                        sections.Add(new_section);
                        section = new_section;
                    }
                    foreach(Section sec in sections){
                        sec.Range.Delete();
                    }
                }
            }
        }
        
        //删除单个表及其标题
        public static void delete_table(Document doc, DocumentBuilder doc_builder, string bookmark, 
            int sec_table_index) {
            doc_builder.MoveToBookmark(bookmark);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, sec_table_index, true);
            table.Range.Delete();
            doc_builder.MoveToBookmark(bookmark);
            Paragraph temp1 = (Paragraph)doc_builder.CurrentParagraph.NextSibling;
            doc_builder.CurrentParagraph.Range.Delete();
            temp1.Range.Delete();
        }

        public static void copy_section(Document doc, string type, int count, int diff) {
            int sec_num = MappingHelper.get_doc_section(type);
            for(int i = 0; i < count; i++) {
                Section new_sec = doc.Sections[sec_num].Clone();
                //ContentFlags.copy += 1;
                copy += 1;
                doc.InsertAfter(new_sec, doc.Sections[sec_num + diff]);
                
            }
        }

        //填写评审组成员信息
        public static void fill_pszcy_info(DocumentBuilder doc_builder, List<string> names) {
            string text = "";
            foreach(string name in names){
                text += name + '、';
                string title = MappingHelper.get_job_title(name);
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark("评审组成员"))
                doc_builder.Write(text);
            names.Insert(0, NamingRules.pingshenzuzhang);
            ContentFlags.pingshenzuchengyuan = names;
        }

        //会议签到表
        public static bool conference_signing(Document doc, DocumentBuilder doc_builder,
            int section_index, int sec_table_index) {
            if(ContentFlags.pingshenzuchengyuan.Count == 0)
                return false;
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int row_index = 1;
            foreach(string name in ContentFlags.pingshenzuchengyuan) {
                if(row_index < ContentFlags.pingshenzuchengyuan.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(1 + row_index, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.hyqdb_name_row, 0);
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.hyqdb_job_row, 0);
                string title = MappingHelper.get_job_title(name);
                doc_builder.Write(title);
                row_index += 1;
            }
            return true;
        }
        
        //会议签到表
        public static bool conference_signing(Document doc, DocumentBuilder doc_builder,
            string bookmark, int sec_table_index) {
            if(ContentFlags.pingshenzuchengyuan.Count == 0)
                return false;
            doc_builder.MoveToBookmark(bookmark);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table,
                sec_table_index, true);
            int row_index = 1;
            foreach(string name in ContentFlags.pingshenzuchengyuan) {
                if(row_index < ContentFlags.pingshenzuchengyuan.Count) {
                    var row = table.Rows[row_index].Clone(true);
                    table.Rows.Insert(1 + row_index, row);
                }
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.hyqdb_name_row, 0);
                doc_builder.Write(name);
                doc_builder.MoveToCell(sec_table_index, row_index, InsertionPos.hyqdb_job_row, 0);
                string title = MappingHelper.get_job_title(name);
                doc_builder.Write(title);
                row_index += 1;
            }
            return true;
        }
        //更新文档域
        public static void update_file(){
            Document doc = new Document(FileConstants.save_root_file);
            if(doc != null){
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
                if(ContentFlags.dmsc_same) {
                    string[] bookmarks = { "调拨6", "调拨7", "配置7"};
                    delete_section(doc, doc_builder, bookmarks);
                    delete_table(doc, doc_builder, "代码审查配置状态报告单", 0);
                }
                if(ContentFlags.dmzc_same) {
                    string[] bookmarks = { "调拨8", "调拨9", "配置12" };
                    delete_section(doc, doc_builder, bookmarks);
                    delete_table(doc, doc_builder, "代码走查配置状态报告单", 0);
                }
                set_file_order(doc, doc_builder);
                set_file_order1(doc, doc_builder);
                NodeCollection nodes = doc.GetChildNodes(NodeType.FieldStart, true);
                foreach(Aspose.Words.Fields.FieldStart field_ref in nodes) {
                    Aspose.Words.Fields.Field field = field_ref.GetField();
                    field.Update();
                }
                doc.Save(FileConstants.save_root_file);
            }
        }

        //编号
        public static void set_file_order(Document doc, DocumentBuilder doc_builder){
           // ContentFlags.set_order();
            Dictionary<string, string> lxwtf_time_dict = ContentFlags.set_lxwtf_order();
            var lxwtf_sort_dict = from objDic in ContentFlags.time_dict1 orderby objDic.Value 
                                      ascending select objDic;//升序
            List<string> new_lxwtf_list = new List<string>();
            List<string> bookmarks = lxwtf_sort_dict.Select(r => r.Key).ToList();
            foreach(string bookmark in lxwtf_time_dict.Keys) {
                if(bookmarks.Contains(bookmark))
                    new_lxwtf_list.Add(lxwtf_time_dict[bookmark]);
            }
            int i = 3;
            foreach(string id_mark in new_lxwtf_list){
                if(doc_builder.MoveToBookmark(id_mark)){
                    if(i < 10)
                        doc_builder.Write("0" + i.ToString());
                    else
                        doc_builder.Write(i.ToString());
                    i += 1;
                }
            }
        }

        //编号
        public static void set_file_order1(Document doc, DocumentBuilder doc_builder) {
            Dictionary<string, List<string>> project_id_dict = ContentFlags.set_order();
            var sort_dict = from objDic in ContentFlags.time_dict2
                                  orderby objDic.Value ascending
                                  select objDic;//升序
            //按顺序排列的时间书签
            List<string> bookmarks = sort_dict.Select(r => r.Key).ToList();

            List<string> new_id_list1 = new List<string>();
            List<string> new_id_list2 = new List<string>();
            List<string> new_id_list3 = new List<string>();
            List<string> new_id_list4 = new List<string>();
            //List<string> bookmarks = lxwtf_sort_dict.Select(r => r.Value).ToList();
            foreach(string bookmark in bookmarks) {
                if(!project_id_dict.ContainsKey(bookmark))
                    continue;
                List<string> positions = project_id_dict[bookmark];
                foreach(string s in positions){
                    if(s.StartsWith("调拨"))
                        new_id_list1.Add(s);
                    else if(s.StartsWith("入库"))
                        new_id_list2.Add(s);
                    else if(s.StartsWith("配置"))
                        new_id_list3.Add(s);
                    else
                        new_id_list4.Add(s);
                }

            }
            int i = 1;
            foreach(string id_mark in new_id_list1) {
                if(doc_builder.MoveToBookmark(id_mark)) {
                    if(i < 10)
                        doc_builder.Write("0" + i.ToString());
                    else
                        doc_builder.Write(i.ToString());
                    i += 1;
                }
            }
            i = 1;
            foreach(string id_mark in new_id_list2) {
                if(doc_builder.MoveToBookmark(id_mark)) {
                    if(i < 10)
                        doc_builder.Write("0" + i.ToString());
                    else
                        doc_builder.Write(i.ToString());
                    i += 1;
                }
            }
            i = 1;
            foreach(string id_mark in new_id_list3) {
                if(doc_builder.MoveToBookmark(id_mark)) {
                    if(i < 10)
                        doc_builder.Write("0" + i.ToString());
                    else
                        doc_builder.Write(i.ToString());
                    i += 1;
                }
            }
            i = 1;
            foreach(string id_mark in new_id_list4) {
                if(doc_builder.MoveToBookmark(id_mark)) {
                    if(i < 10)
                        doc_builder.Write("0" + i.ToString());
                    else
                        doc_builder.Write(i.ToString());
                    i += 1;
                }
            }
        }

    }
}
