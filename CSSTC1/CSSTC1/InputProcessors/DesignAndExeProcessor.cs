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
            Document doc = new Document(FilePaths.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);

            //第五章 文档审查结果记录入库时间
            if(TimeStamp.wdscqr_time == null){
                //MessageBox.Show("未输入文档审查确认时间");
                return false;
            }
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
            doc.Save(FilePaths.save_root_file);
            return true;

        }
    }
}
