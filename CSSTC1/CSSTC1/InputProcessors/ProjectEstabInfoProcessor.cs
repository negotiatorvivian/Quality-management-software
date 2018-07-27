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

        public void fill_estab_info(string lx_time, List<string> lq_times, string pl_time){
            int lixiang_time = int.Parse(lx_time);
            Document doc = new Document(FilePaths.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            Bookmark mark = doc.Range.Bookmarks["项目开始时间"];
            string xmks_time = "";
            if(mark != null){
                xmks_time = mark.Text;
                string[] items = xmks_time.Split('/');
                DateTime start_date = new DateTime(int.Parse(items[0]), int.Parse(items[1]), int.Parse(items[2]));
                string contact_time = DateTime.Parse(start_date.ToString("yyyy-MM-dd")).
                    AddDays(0 - lixiang_time).ToShortDateString(); 
                contact_time.Replace('/', '-');
                if(doc_builder.MoveToBookmark("联系时间")){
                    doc_builder.Write(contact_time);
                    doc.Save(FilePaths.save_root_file);
                }
            }
            if(doc_builder.MoveToBookmark("偏离联系时间")){
                if(pl_time.Length > 0){
                    doc_builder.Write(pl_time);
                }
                else{
                    Bookmark mark1 = doc.Range.Bookmarks["合同偏离通知单"];
                    mark1.Remove();
                }
                doc.Save(FilePaths.save_root_file);

            }
        }
    }
}
