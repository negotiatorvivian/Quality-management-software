using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace CSSTC1.CommonUtils {
    class DateHelper {
        public static Document fill_time_blank(Document doc, DocumentBuilder doc_builder,
    string bookmark, string origin_time, int diff) {
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                        AddDays(0 - diff).ToLongDateString();
            if(doc_builder.MoveToBookmark(bookmark)) {
                doc_builder.Write(temp);
            }
            return doc;
        }

        public static string cal_time(string origin_time, int diff) {
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                        AddDays(0 - diff).ToShortDateString();
            return temp;
        }

        public static DateTime cal_date(string origin_time, int diff) {
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            DateTime temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).AddDays(0 - diff);
            return temp;
        }
    }
}
