using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;
using System.Windows;

namespace CSSTC1.InputProcessors {
    class DesignAndExeProcessor {
        public bool fill_time_line(string pl_time){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            //测试就绪评审环节是否有问题
            if(ContentFlags.pianli_3 == 0){
                string[] marks = { "合同偏离通知单2" };
                OperationHelper.delete_section(doc, doc_builder, marks);
            }

            if(ContentFlags.pianli_4 == 0) {
                string[] marks = { "合同偏离通知单3" };
                OperationHelper.delete_section(doc, doc_builder, marks);

            }
            else{
                if(doc_builder.MoveToBookmark("偏离联系时间3"))
                    doc_builder.Write(pl_time);

            }
            string csjxps_time1 = TimeStamp.csjxps_time[0];
            
            if(doc_builder.MoveToBookmark("联系委托方第七次")){
                DateTime t1 = DateHelper.cal_date(csjxps_time1, 3);
                ContentFlags.time_dict1.Add("联系委托方第七次", t1);
                doc_builder.Write(t1.ToLongDateString());
            }
            
            if(doc_builder.MoveToBookmark("测试就绪内部评审时间")){
                doc_builder.Write(TimeStamp.csjxps_format_time[0]);
                DateTime t2 = DateHelper.cal_date(TimeStamp.csjxps_time[0], 0);
                ContentFlags.time_dict2.Add("测试就绪内部评审时间", t2);
            }
            
            if(doc_builder.MoveToBookmark("搭建环境被测件接收时间")){
                DateTime t1 = DateHelper.cal_date(csjxps_time1, 1);
                ContentFlags.time_dict2.Add("搭建环境被测件接收时间", t1);
                doc_builder.Write(t1.ToLongDateString());
            }
            if(TimeStamp.csjxps_time.Count > 1){
                string csjxps_time2 = TimeStamp.csjxps_time[1];
                if(doc_builder.MoveToBookmark("测试就绪内部评审时间1"))
                    doc_builder.Write(TimeStamp.csjxps_format_time[1]);
                if(doc_builder.MoveToBookmark("搭建环境被测件接收时间1")){
                    DateTime t1 = DateHelper.cal_date(csjxps_time2, 3);
                    ContentFlags.time_dict2.Add("搭建环境被测件接收时间1", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }

            }
            //测试环境标识编号
            this.change_file_structure(doc, doc_builder);
            if(ContentFlags.jingtaifenxi > 0 && doc_builder.MoveToBookmark("静态分析测试环境标识")) {
                doc_builder.Write((ContentFlags.test_env_id + 1).ToString());
                ContentFlags.test_env_id += 1;
            }
            if(ContentFlags.daimashencha > 0 && doc_builder.MoveToBookmark("代码审查测试环境标识")) {
                doc_builder.Write((ContentFlags.test_env_id + 1).ToString());
                ContentFlags.test_env_id += 1;
            }
            if(ContentFlags.daimazoucha > 0 && doc_builder.MoveToBookmark("代码走查测试环境标识")) {
                doc_builder.Write((ContentFlags.test_env_id + 1).ToString());
                ContentFlags.test_env_id += 1;
            }
            if(doc_builder.MoveToBookmark("动态测试环境标识")) {
                doc_builder.Write((ContentFlags.test_env_id + 1).ToString());
                ContentFlags.test_env_id += 1;
            }
            if(ContentFlags.peizhiceshi > 0){
               // string pzcs_lxwtf_time2 = DateHelper.cal_time(TimeStamp.hgdtcs_time, 7);
                //string pzcs_hg_time = TimeStamp.hgdtcs_format_time;
                if(doc_builder.MoveToBookmark("联系委托方第八次")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.hgdtcs_time, 14);
                    ContentFlags.time_dict1.Add("联系委托方第八次", t1);
                    ContentFlags.time_dict2.Add("联系委托方第八次", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }
                if(doc_builder.MoveToBookmark("联系委托方第九次")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.hgdtcs_time, 7);
                    ContentFlags.time_dict1.Add("联系委托方第九次", t1);
                    doc_builder.Write(t1.ToLongDateString());
                    
                }
                if(doc_builder.MoveToBookmark("配置项动态回归被测件接收时间")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.hgdtcs_time, 3);
                    ContentFlags.time_dict2.Add("配置项动态回归被测件接收时间", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }
                if(doc_builder.MoveToBookmark("回归测试时间")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.hgdtcs_time, 0);
                    ContentFlags.time_dict2.Add("回归测试时间", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }
            } 
            if(ContentFlags.luojiceshi > 0){
                string ljcs_lxwtf_time1 = TimeStamp.ljcs_format_time;
                if(doc_builder.MoveToBookmark("逻辑测试时间"))
                    doc_builder.Write(ljcs_lxwtf_time1);
            }
            if(ContentFlags.xitongceshi > 0){
                if(doc_builder.MoveToBookmark("联系委托方第十次")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.slxtcs_time, 3);
                    ContentFlags.time_dict1.Add("联系委托方第十次", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }
                if(doc_builder.MoveToBookmark("系统测试时间")){
                    DateTime t1 = DateHelper.cal_date(TimeStamp.slxtcs_time, 0);
                    ContentFlags.time_dict2.Add("系统测试时间", t1);
                    doc_builder.Write(t1.ToLongDateString());
                }
                if(ContentFlags.xitonghuiguiceshi > 0){
                    string xthgcs_time = DateHelper.cal_time(TimeStamp.xthgcs_time, 3);
                    if(doc_builder.MoveToBookmark("联系委托方第十一次")){
                        DateTime t1 = DateHelper.cal_date(TimeStamp.xthgcs_time, 3);
                        ContentFlags.time_dict1.Add("联系委托方第十一次", t1);
                        doc_builder.Write(t1.ToLongDateString());
                    }
                    if(doc_builder.MoveToBookmark("系统回归测试时间")){
                        DateTime t1 = DateHelper.cal_date(TimeStamp.xthgcs_time, 0);
                        ContentFlags.time_dict2.Add("系统回归测试时间", t1);
                        doc_builder.Write(t1.ToLongDateString());
                    }
                }
            }
            doc.Save(FileConstants.save_root_file);
            return true;

        }

        public void change_file_structure(Document doc, DocumentBuilder doc_builder){
            if(TimeStamp.csjxps_time.Count == 0 || ContentFlags.test_envs.Count == 0){            
                OperationHelper.delete_section(doc, doc_builder, "测试就绪第二次", "ceshijiuxu2");
                ContentFlags.ceshijiuxu2 = 0;
            }
        }
    }
}
