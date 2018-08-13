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
        public void fill_estab_info(string lx_time, List<string> lq_times, string pl_time){
            int lixiang_time = int.Parse(lx_time);
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            string xmks_time = TimeStamp.xiangmukaishishijian ;
            if(xmks_time.Length > 0) {
                DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第一次", xmks_time, lixiang_time);
                DateHelper.fill_time_blank(doc, doc_builder, "任务通知时间", xmks_time, lixiang_time - 1);
                DateHelper.fill_time_blank(doc, doc_builder, "内部评审时间", xmks_time, lixiang_time - 4);
                DateHelper.fill_time_blank(doc, doc_builder, "联系委托方第二次", xmks_time, lixiang_time - 6);
                //被测件接收时间以默认为准还是输入为准？？
                //doc = this.fill_time_blank(doc, doc_builder, "被测件接收时间", xmks_time, lixiang_time - 7);
            }
            //else{
            //    doc_builder.MoveToBookmark("被测件调拨2");
            //    doc_builder.CurrentSection.Range.Delete();
            //    doc_builder.MoveToBookmark("被测件调拨2配置状态");
            //    doc_builder.CurrentSection.Range.Delete();

            //}
            this.change_file_struc(doc, doc_builder, "被测件领取次数");
            for(int i = 0; i < lq_times.Count; i ++){
                if(doc_builder.MoveToBookmark("入库日期" + (i + 1).ToString()))
                    doc_builder.Write(lq_times[i]);
            }
            if(doc_builder.MoveToBookmark("偏离联系时间")){
                if(pl_time.Length > 0){
                    string[] temp = pl_time.Split('/');
                    string temp1 = temp[0] + "年" + temp[1] + "月" + temp[2] + "日";
                    doc_builder.Write(temp1);
                }
                else{
                    ContentFlags.pianli_1 = 0;
                    string[] marks = {"合同偏离通知单"};
                    OperationHelper.delete_section(doc, doc_builder,marks);
                }
            }
            //NodeCollection nodes = doc.GetChildNodes(NodeType.FieldStart, true);
            //foreach(Aspose.Words.Fields.FieldStart field_ref in nodes) {
            //    Aspose.Words.Fields.Field field = field_ref.GetField();
            //    field.Update();
            //}
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
