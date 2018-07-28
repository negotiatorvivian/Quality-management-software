using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;

namespace CSSTC1.InputProcessors {
    public class ProjectEstabInfoProcessor {
        
        public void show_estab_info(){
            Globals.ThisDocument.project_estab_info.Show();
        }

        //填写项目时间线
        public void fill_estab_info(string lx_time, List<string> lq_times, string pl_time){
            int lixiang_time = int.Parse(lx_time);
            ContentFlags.lingqushijian = lq_times;
            Document doc = new Document(FilePaths.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Bookmark mark = doc.Range.Bookmarks["项目开始时间"];
            string xmks_time = "";
            if(mark != null) {
                xmks_time = mark.Text;
                ContentFlags.xiangmukaishishijian = xmks_time;
                doc = fill_time_blank(doc, doc_builder, "联系委托方第一次", xmks_time, lixiang_time);
                doc = fill_time_blank(doc, doc_builder, "任务通知时间", xmks_time, lixiang_time - 1);
                doc = fill_time_blank(doc, doc_builder, "内部评审时间", xmks_time, lixiang_time - 4);
                doc = fill_time_blank(doc, doc_builder, "联系委托方第二次", xmks_time, lixiang_time - 6);
                //被测件接收时间以默认为准还是输入为准？？
                //doc = this.fill_time_blank(doc, doc_builder, "被测件接收时间", xmks_time, lixiang_time - 7);
            }
            //if(lq_times.Count > 1){
            //    for(int i = 1; i < lq_times.Count; i++){
            //        string bcjjs_time = lq_times[i];
            //        doc = fill_time_blank(doc, doc_builder, "被测件接收时间" + (i + 1).ToString(), 
            //            bcjjs_time, 0);
            //    }
            //}
            else{
                doc_builder.MoveToBookmark("被测件调拨2");
                doc_builder.CurrentSection.Range.Delete();
                doc_builder.MoveToBookmark("被测件调拨2配置状态");
                doc_builder.CurrentSection.Range.Delete();

            }
            if(doc_builder.MoveToBookmark("偏离联系时间")){
                if(pl_time.Length > 0){
                    doc_builder.Write(pl_time);
                }
                else{
                    doc_builder.MoveToBookmark("合同偏离通知单");
                    doc_builder.CurrentSection.Range.Delete();
                }
            }
            NodeCollection nodes = doc.GetChildNodes(NodeType.FieldStart, true);
            foreach(Aspose.Words.Fields.FieldStart field_ref in nodes) {
                Aspose.Words.Fields.Field field = field_ref.GetField();
                field.Update();
            }
            doc.Save(FilePaths.save_root_file);
        }

        public static Document fill_time_blank(Document doc, DocumentBuilder doc_builder, 
            string bookmark, string origin_time, int diff ){
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                        AddDays(0 - diff).ToShortDateString();
            if(doc_builder.MoveToBookmark(bookmark)) {
                doc_builder.Write(temp);
            }
            //doc.Save(FilePaths.save_root_file);
            return doc;
        }

        public static string cal_time(string origin_time, int diff ){
            string[] items = origin_time.Split('/');
            DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
            string temp = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                        AddDays(0 - diff).ToShortDateString();
            return temp;
    }
    }
}
