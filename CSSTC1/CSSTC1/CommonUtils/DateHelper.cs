using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;

namespace CSSTC1.CommonUtils {
    class DateHelper {
        public static void fill_time_blank(Document doc, DocumentBuilder doc_builder, string bookmark, 
            string origin_time, int diff) {
            if(origin_time == null || origin_time.Length == 0){
                string temp =  new DateTime().ToLongDateString();
                if(doc_builder.MoveToBookmark(bookmark)) {
                    doc_builder.Write(temp);
            }
            }
            else{
                string[] items = origin_time.Split('/');
                DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
                string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                            AddDays(0 - diff).ToLongDateString();
                if(doc_builder.MoveToBookmark(bookmark)) {
                    doc_builder.Write(temp);
                }
            }
        }

        public static string cal_time(string origin_time, int diff) {
            if(origin_time == null || origin_time.Length == 0)
                return new DateTime().ToShortDateString();
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                        AddDays(0 - diff).ToLongDateString();
            return temp;
        }

        public static DateTime cal_date(string origin_time, int diff) {
            if(origin_time == null || origin_time.Length == 0)
                return new DateTime();
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            DateTime temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).AddDays(0 - diff);
            return temp;
        }

        public static string cal_general_time(string date_str){
            string datetime = "";
            int index1 = date_str.IndexOf("月");
            string temp1 = date_str.Substring(0, index1 + 1);
            Regex reg = new Regex("[0-9]+");
            Match match = reg.Match(date_str);
            if(match.Groups != null) {
                string date = match.Groups[match.Groups.Count - 1].ToString();
                int date_num = int.Parse(date);
                if(date_num < 11)
                    datetime = "上旬";
                else if(date_num < 21)
                    datetime = "中旬";
                else
                    datetime = "下旬";
            }
            temp1 = temp1 + datetime;
            return temp1;
        }
    }
}
