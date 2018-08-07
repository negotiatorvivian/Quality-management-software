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
                    //doc.Save(FileConstants.save_root_file);
                }
            }
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

        public static void update_file(){
            Document doc = new Document(FileConstants.save_root_file);
            if(doc != null){
                DocumentBuilder doc_builder = new DocumentBuilder(doc);
                NodeCollection nodes = doc.GetChildNodes(NodeType.FieldStart, true);
                foreach(Aspose.Words.Fields.FieldStart field_ref in nodes) {
                    Aspose.Words.Fields.Field field = field_ref.GetField();
                    field.Update();
                }
                doc.Save(FileConstants.save_root_file);
            }
        }
    }
}
