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
        public bool fill_time_line(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            //测试就绪评审环节是否有问题
            if(ContentFlags.pianli_3 == 0){
                string[] marks = { "合同偏离通知单2" };
                OperationHelper.delete_section(doc, doc_builder, marks);
            }
            string csjxps_time1 = TimeStamp.csjxps_time[0];
            string lxqtf_time3 = DateHelper.cal_time(csjxps_time1, 3);
            string bcjjs_time = DateHelper.cal_time(csjxps_time1, 1);

            if(doc_builder.MoveToBookmark("联系委托方第七次"))
                doc_builder.Write(lxqtf_time3);
            if(doc_builder.MoveToBookmark("测试就绪内部评审时间"))
                doc_builder.Write(TimeStamp.csjxps_format_time[0]);
            if(doc_builder.MoveToBookmark("搭建环境被测件接收时间"))
                doc_builder.Write(bcjjs_time);
            if(TimeStamp.csjxps_time.Count > 1){
                string csjxps_time2 = TimeStamp.csjxps_time[1];
                string bcjjs_time2 = DateHelper.cal_time(csjxps_time2, 3);
                if(doc_builder.MoveToBookmark("测试就绪内部评审时间1"))
                    doc_builder.Write(TimeStamp.csjxps_format_time[1]);
                if(doc_builder.MoveToBookmark("搭建环境被测件接收时间1"))
                    doc_builder.Write(bcjjs_time2);
            }
            this.change_file_structure(doc, doc_builder);

            if(ContentFlags.peizhiceshi > 0){
                string pzcs_lxwtf_time1 = DateHelper.cal_time(TimeStamp.hgdtcs_time, 14);
                string pzcs_lxwtf_time2 = DateHelper.cal_time(TimeStamp.hgdtcs_time, 7);
                string pzcs_bcjjs_time1 = DateHelper.cal_time(TimeStamp.hgdtcs_time, 3);
                string pzcs_hg_time = TimeStamp.hgdtcs_format_time;
                if(doc_builder.MoveToBookmark("联系委托方第八次"))
                    doc_builder.Write(pzcs_lxwtf_time1);
                if(doc_builder.MoveToBookmark("联系委托方第九次"))
                    doc_builder.Write(pzcs_lxwtf_time2);
                if(doc_builder.MoveToBookmark("配置项动态回归被测件接收时间"))
                    doc_builder.Write(pzcs_bcjjs_time1);
                if(doc_builder.MoveToBookmark("回归测试时间"))
                    doc_builder.Write(pzcs_hg_time);
            } 
            if(ContentFlags.luojiceshi > 0){
                string ljcs_lxwtf_time1 = TimeStamp.ljcs_format_time;
                if(doc_builder.MoveToBookmark("逻辑测试时间"))
                    doc_builder.Write(ljcs_lxwtf_time1);
            }
            if(ContentFlags.xitongceshi > 0){
                string xtcs_lxwtf_time = DateHelper.cal_time(TimeStamp.slxtcs_time, 3);
                if(doc_builder.MoveToBookmark("联系委托方第十次"))
                    doc_builder.Write(xtcs_lxwtf_time);
                if(doc_builder.MoveToBookmark("系统测试时间"))
                    doc_builder.Write(TimeStamp.slxtcs_format_time);
                if(ContentFlags.xitonghuiguiceshi > 0){
                    string xthgcs_time = DateHelper.cal_time(TimeStamp.xthgcs_time, 3);
                    if(doc_builder.MoveToBookmark("联系委托方第十一次"))
                        doc_builder.Write(xthgcs_time);
                    if(doc_builder.MoveToBookmark("系统回归测试时间"))
                        doc_builder.Write(TimeStamp.xthgcs_format_time);
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
