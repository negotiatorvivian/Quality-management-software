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

            //第五章 文档审查结果记录入库时间
            if(TimeStamp.wdscqr_time == null){
                //MessageBox.Show("未输入文档审查确认时间");
                return false;
            }
            //是否有测试就绪评审环节
            if(pl_time.Length > 0){
                if(ContentFlags.pianli_3 == 0){
                    string[] marks = { "合同偏离通知单2" };
                    OperationHelper.delete_section(doc, doc_builder, marks);
                }
            }
            //else{
            //    if(doc_builder.MoveToBookmark("测试就绪内部评审时间"))
            //        doc_builder.Write(pl_time);
            //}
            DateHelper.fill_time_blank(doc, doc_builder, "文档审查结果记录入库时间", TimeStamp.wdscqr_time, 1);
            //文档审查确认时间
            if(doc_builder.MoveToBookmark("文档审查确认时间"))
                doc_builder.Write(TimeStamp.wdscqr_format_time);
            //文档审查回归时间
            if(doc_builder.MoveToBookmark("文档审查回归时间静态分析审查时间"))
                doc_builder.Write(TimeStamp.wdschg_format_time);

            //联系研制方提供被测软件源代码时间
            DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第五次", TimeStamp.jtfxsc_time, 7);
            //静态分析审查时间
            if(doc_builder.MoveToBookmark("静态分析审查时间"))
                doc_builder.Write(TimeStamp.jtfxsc_format_time);
            //静态分析结果入库时间
            DateHelper.fill_time_blank(doc, doc_builder, "静态分析结果入库时间", TimeStamp.jtfxqr_time, 1);
            //联系委托方第六次:研制方来测评中心确认静态分析问题时间
            DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第六次", TimeStamp.jtfxqr_time, 3);
            if(doc_builder.MoveToBookmark("静态分析回归时间"))
                doc_builder.Write(TimeStamp.jtfxhg_format_time);
            //测试说明内部评审时间
            if(doc_builder.MoveToBookmark("测试说明内部评审时间"))
                doc_builder.Write(TimeStamp.cssmps_format_time);
            doc.Save(FileConstants.save_root_file);
            return true;

        }
    }
}
