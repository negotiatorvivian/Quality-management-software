using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using CSSTC1.ConstantVariables;
using CSSTC1.CommonUtils;
using System.Windows.Forms;

namespace CSSTC1.InputProcessors {
    public class DemandAnaProcessor {
        //填写第三四部分项目时间
        public void fill_time_info(string ns_time, string pl_time){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            if(doc_builder.MoveToBookmark("偏离联系时间2") && pl_time.Length > 0){
                doc_builder.Write(pl_time);
                DateTime t = DateHelper.cal_date(TimeStamp.pianli2_time, 0);
                ContentFlags.time_dict1.Add("偏离联系时间2", t);
            }
            //没有偏离，删除偏离表格
            else if(pl_time.Length == 0){
                string[] marks = { "合同偏离通知单1" };
                OperationHelper.delete_section(doc, doc_builder, marks);
            }
            //填写测试大纲部分
            if(ContentFlags.ceshidagang) {
                this.fill_csdg_basic_time(doc, doc_builder, ns_time);
            }
            //填写需求说明部分
            else {
                this.fill_xqsm_basic_time(doc, doc_builder, ns_time);
                DateTime t = DateHelper.cal_date(ns_time, 0);
                ContentFlags.time_dict2.Add("测试说明与计划内审时间", t);
            }
        }

        public void fill_csdg_basic_time(Document doc, DocumentBuilder doc_builder, string ns_time){
            if(TimeStamp.lingqushijian.Count > 0) {
                string temp = TimeStamp.lingqushijian[TimeStamp.lingqushijian.Count - 1];
                DateHelper.fill_time_blank(doc, doc_builder, "质量保证计划日期", temp, -7);

                string zhiliangbaozheng = DateHelper.cal_date(temp, -7).ToLongDateString();
                int index1 = zhiliangbaozheng.IndexOf("月");
                zhiliangbaozheng = zhiliangbaozheng.Substring(0, index1 + 1);
                if(doc_builder.MoveToBookmark("质量保证计划时间"))
                    doc_builder.Write(zhiliangbaozheng);
                //软件测试过程需求及需求变更评审表???
                DateHelper.fill_time_blank(doc, doc_builder, "需求及需求变更评审", temp, 0);//
                if(doc_builder.MoveToBookmark("大纲内审时间"))
                    doc_builder.Write(ns_time);
                string ceshidagang_shencha = TimeStamp.ceshisc_time;
                //DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第三次", );
                DateTime t1 = DateHelper.cal_date(ceshidagang_shencha, 7);
                DateTime t2 = DateHelper.cal_date(ceshidagang_shencha, -1);
                if(doc_builder.MoveToBookmark("联系委托方第三次"))
                    doc_builder.Write(t1.ToLongDateString());
                if(doc_builder.MoveToBookmark("测试大纲出库时间"))
                    doc_builder.Write(t2.ToLongDateString());
                //DateHelper.fill_time_blank(doc, doc_builder, "测试大纲出库时间", ceshidagang_shencha, -1);
                string ceshidagang_shencha1 = TimeStamp.ceshisc_format_time;
                int index = ceshidagang_shencha1.IndexOf("月");
                string temp1 = ceshidagang_shencha1.Substring(0, index + 1);
                if(doc_builder.MoveToBookmark("测试大纲审查时间"))
                    doc_builder.Write(temp1);
                doc.Save(FileConstants.save_root_file);
                ContentFlags.time_dict1.Add("联系委托方第三次", t1);
                ContentFlags.time_dict2.Add("测试大纲出库时间", t2);
            }
            else
                MessageBox.Show("请先填写项目立项阶段信息");
        }

        public void fill_xqsm_basic_time(Document doc, DocumentBuilder doc_builder, string ns_time){
            if(doc_builder.MoveToBookmark("测试说明与计划内审时间"))
                doc_builder.Write(ns_time);
            if(TimeStamp.ceshisc_time != null){
                DateTime t = DateHelper.cal_date(TimeStamp.ceshisc_time, 7);
                ContentFlags.time_dict1.Add("联系委托方第四次", t);
                if(doc_builder.MoveToBookmark("联系委托方第四次"))
                    doc_builder.Write(t.ToLongDateString());
                if(doc_builder.MoveToBookmark("测试说明与计划审查时间"))
                    doc_builder.Write(TimeStamp.ceshisc_format_time);
                DateTime t1 = DateHelper.cal_date(TimeStamp.ceshisc_time, -1);
                ContentFlags.time_dict2.Add("测试说明与计划出库时间", t);
                if(doc_builder.MoveToBookmark("测试说明与计划出库时间"))
                    doc_builder.Write(t1.ToLongDateString());
            }
            doc.Save(FileConstants.save_root_file);
        }
    }
}
