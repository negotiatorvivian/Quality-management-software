using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.FileProcessors.models;

namespace CSSTC1.CommonUtils {
    class OperationHelper {
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

        public static void copy_section(Document doc, string type, int count, int diff) {
            int sec_num = MappingHelper.get_doc_section(type);
            for(int i = 0; i < count; i++) {
                Section new_sec = doc.Sections[sec_num].Clone();
                ContentFlags.copy += 1;
                doc.InsertAfter(new_sec, doc.Sections[sec_num + diff]);
                
            }
        }

       
    }
}
