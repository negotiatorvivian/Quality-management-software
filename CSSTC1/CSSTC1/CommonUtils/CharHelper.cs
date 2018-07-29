using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace CSSTC1.CommonUtils {
    class CharHelper {
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
    }
}
