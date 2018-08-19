using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;

namespace CSSTC1.InputProcessors {
    public class ProjectEstabInfoProcessor {
        
        public void show_estab_info(){
            Globals.ThisDocument.project_estab_info.Show();
        }

        //填写项目时间线
        public void fill_estab_info(string lx_time, List<DateTime> lq_times, DateTime pl_time) {
            int lixiang_time = int.Parse(lx_time);
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            List<string> keys = new List<string>();
            List<string> values = new List<string>();
            string xmks_time = TimeStamp.xiangmukaishishijian ;
            if(xmks_time.Length > 0) {
                DateTime lxwtf_time1 = DateHelper.cal_date(xmks_time, lixiang_time);
                keys.Add("联系委托方第一次");
                values.Add(lxwtf_time1.ToLongDateString());
                ContentFlags.time_dict1.Add("联系委托方第一次", lxwtf_time1);
                string rwtzd_time = DateHelper.cal_time(xmks_time, lixiang_time - 1);
                keys.Add("任务通知时间");
                values.Add(rwtzd_time);
                DateTime ns_time = DateHelper.cal_date(xmks_time, lixiang_time - 1);
                keys.Add("内部评审时间");
                values.Add(ns_time.ToLongDateString());
                ContentFlags.time_dict2.Add("内部评审时间", ns_time);
                DateTime lxwtf_time2 = DateHelper.cal_date(xmks_time, lixiang_time);
                keys.Add("联系委托方第二次");
                values.Add(lxwtf_time2.ToLongDateString());
                ContentFlags.time_dict1.Add("联系委托方第二次", lxwtf_time2);
                //被测件接收时间以默认为准还是输入为准？？
                //doc = this.fill_time_blank(doc, doc_builder, "被测件接收时间", xmks_time, lixiang_time - 7);
            }
            //else{
            //    doc_builder.MoveToBookmark("被测件调拨2");
            //    doc_builder.CurrentSection.Range.Delete();
            //    doc_builder.MoveToBookmark("被测件调拨2配置状态");
            //    doc_builder.CurrentSection.Range.Delete();

            //}
            string csdgps_time = TimeStamp.ceshisc_format_time;
            string datetime = DateHelper.cal_general_time(csdgps_time);
            keys.Add("测试大纲评审时间");
            values.Add(datetime);

            this.change_file_struc(doc, doc_builder, "被测件领取次数");
            for(int i = 0; i < lq_times.Count; i ++){
                keys.Add("入库日期" + (i + 1).ToString());
                values.Add(lq_times[i].ToLongDateString());
                ContentFlags.time_dict2.Add("入库日期" + (i + 1).ToString(), lq_times[i]);
                
                string rk_general_time = DateHelper.cal_general_time(lq_times[i].ToLongDateString());
                keys.Add("入库时间" + (i + 1).ToString() + "精确到月");
                values.Add(rk_general_time);
            }
          
            if(pl_time != DateTime.MaxValue){
                keys.Add("偏离联系时间");
                values.Add(pl_time.ToLongDateString());
                ContentFlags.time_dict1.Add("偏离联系时间", pl_time);
            }
            else{
                ContentFlags.pianli_1 = 0;
                string[] marks = {"合同偏离通知单"};
                OperationHelper.delete_section(doc, doc_builder,marks);
            }
            int index = 0;
            foreach(string mark in keys){
                if(doc_builder.MoveToBookmark(mark))
                    doc_builder.Write(values[index]);
                index += 1;
            }

            doc.Save(FileConstants.save_root_file);
        }

        public void change_file_struc(Document doc, DocumentBuilder doc_builder, string content){
            switch(content){
                case "被测件领取次数":{
                    if(ContentFlags.lingqucishu == 1) {
                        string[] marks = { "被测件调拨2", "被测件调拨2配置状态" };
                        OperationHelper.delete_section(doc, doc_builder, marks);
                        doc.Save(FileConstants.save_root_file);
                    }
                    if(ContentFlags.lingqucishu > 2) {
                        string mark = "被测件清单";
                        OperationHelper.copy_section(doc, mark, ContentFlags.lingqucishu - 2, 1);
                        string mark1 = "配置报告单";
                        OperationHelper.copy_section(doc, mark1, ContentFlags.lingqucishu - 2, 1);
                       // doc.Save(FilePaths.save_root_file);
                    }
                    break;
                }
                default:break;
        }
        }
    }
}
