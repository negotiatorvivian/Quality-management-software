using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;
using Aspose.Words.Tables;
using CSSTC1.ConstantVariables;
using System.Text.RegularExpressions;
using Aspose.Words;
using CSSTC1.CommonUtils;
using System.Windows;

namespace CSSTC1.FileProcessors.writers.BuildEnvironment {
    class FileWriter9 {
        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict;
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict;
        List<List<string>> files_ids = new List<List<string>>();
        List<SoftwareItems> softwares = new List<SoftwareItems>();
        public string software_names = "";
        List<int> time_diff = new List<int>();
        private TestEnvChartHelper helper;
        public FileWriter9(Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string,
            List<DynamicHardwareItems>> yingjianpeizhi_dict) {
            this.ruanjianpeizhi_dict = ruanjianpeizhi_dict;
            this.yingjianpeizhi_dict = yingjianpeizhi_dict;
            time_diff.Add(ContentFlags.lingqucishu * 2);
            time_diff.Add(ContentFlags.pianli_1);
            time_diff.Add(ContentFlags.pianli_2);
            time_diff.Add(ContentFlags.wendangshencha);
            time_diff.Add(ContentFlags.jingtaifenxi);
            time_diff.Add(ContentFlags.daimashencha);
            time_diff.Add(ContentFlags.daimazoucha);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.pianli_3);
            helper = new TestEnvChartHelper();
            bool res = this.write_charts();
            if(!res)
                MessageBox.Show("没有软件选择两次测试环境！");
        }
        public bool write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            if(TimeStamp.csjxps_time.Count == 0)
                return false;
            helper.write_bcjqd_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict,
                InsertionPos.djhj_bcjqd_section1, InsertionPos.djhj_bcjqd_sec_table, 
                InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row, this.time_diff);
            helper.write_bcjdbd_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict,
                InsertionPos.djhj_bcjqd_section1, InsertionPos.djhj_bcjdbd_sec_table,
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            Dictionary<string, List<SoftwareItems>> temp1 = new Dictionary<string,List<SoftwareItems>>();
            Dictionary<string, List<DynamicHardwareItems>> temp2 = new Dictionary
                <string,List<DynamicHardwareItems>>();
            List<string> softwares = this.ruanjianpeizhi_dict.Select(r => r.Key).ToList();
            foreach(TestEnvironment test_env in ContentFlags.test_envs){
                
                if(test_env.ceshi_didian2.Equals("本测评中心")){
                    string software_name = test_env.rj_mingcheng;
                    if(this.ruanjianpeizhi_dict.ContainsKey(software_name))
                        temp1.Add(software_name, this.ruanjianpeizhi_dict[software_name]);

                    if(this.yingjianpeizhi_dict.ContainsKey(software_name))
                        temp2.Add(software_name, this.yingjianpeizhi_dict[software_name]);
                    
                }
            }
            if(temp1.Count == 0)
                this.delete_table(doc, doc_builder, "搭建环境被测件领取清单");
            helper.write_bcjlqqd_chart(doc, doc_builder, temp1, temp2, InsertionPos.djhj_bcjqd_section1, 
                InsertionPos.djhj_bcjlqqd_sec_table, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, InsertionPos.djhj_bcjlqqd_type_row, this.time_diff);
            this.write_pzztbbd_chart(doc, doc_builder, "测试就绪软件名称2");
            helper.write_csgjhsb_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.djhj_cshjqr_section1, InsertionPos.djhj_cshjqr_sec_table, this.time_diff);
            helper.write_csgjhsbhcd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.djhj_cshjhc_section1, InsertionPos.djhj_cshjhc_sec_table, 1, this.time_diff);
            ContentFlags.beiceruanjianshuliang1 = ruanjianpeizhi_dict.Count;
            this.time_diff.Add(ContentFlags.beiceruanjianshuliang1 * 2);
            //this.merge_location(doc, doc_builder, "搭建环境配置状态报告单1");
            int index = InsertionPos.djhj_hyqdb_section1;
            foreach(int i in time_diff)
                index += i;
            OperationHelper.conference_signing(doc, doc_builder, index, InsertionPos.djhj_hyqdb_sec_table);
            doc.Save(FileConstants.save_root_file);
            return true;
        }

        public void delete_table(Document doc, DocumentBuilder doc_builder, string bookmark){
            doc_builder.MoveToBookmark(bookmark);
            Table table = (Table)doc_builder.CurrentSection.GetChild(NodeType.Table, 2, true);
            table.Range.Delete();
            doc_builder.MoveToBookmark(bookmark);
            Paragraph temp1 = (Paragraph)doc_builder.CurrentParagraph.NextSibling;
            doc_builder.CurrentParagraph.Range.Delete();  
            temp1.Range.Delete();
            //doc_buiilder.CurrentNode.
        }

        public void write_pzztbbd_chart(Document doc, DocumentBuilder doc_builder, string bookmark) {
            string text = "";
            foreach(TestEnvironment test_env in ContentFlags.test_envs){
                text += test_env.rj_mingcheng + '、';
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(text);
        }

        public bool merge_location(Document doc, DocumentBuilder doc_builder, string bookmark){
            if(!doc_builder.MoveToBookmark(bookmark))
                return false;
            Dictionary<string, List<string>> software_test_envs = new Dictionary<string,List<string>>();
            foreach(TestEnvironment test_env in ContentFlags.test_envs){
                if(!software_test_envs.ContainsKey(test_env.ceshi_didian2)){
                    List<string> locations = new List<string>();
                    locations.Add(test_env.rj_mingcheng);
                    software_test_envs.Add(test_env.ceshi_didian2, locations);
                }
                else{
                    software_test_envs[test_env.ceshi_didian2].Add(test_env.rj_mingcheng);
                }
            }
            //string text = "";
            doc_builder.MoveToBookmark(bookmark);
            foreach(string location in software_test_envs.Keys){
                Paragraph old_para = (Paragraph)doc_builder.CurrentParagraph;
                int t = doc_builder.CurrentSection.IndexOf(old_para);
                old_para = (Paragraph)doc_builder.CurrentSection.GetChild(NodeType.Paragraph, t, true);
                Paragraph para = (Paragraph)doc_builder.CurrentParagraph.Clone(true);
                Paragraph temp = new Paragraph(doc);
                
                doc.InsertAfter(para, old_para);
                doc_builder.MoveTo(old_para);
                string names = "";
                foreach(string software in software_test_envs[location]){
                    names += software + '、';
                }
                names = names.Substring(0, names.Length - 1);
                names += "动态测试环境已搭建，位于" + location;
                doc_builder.Write(names);
                doc_builder.MoveTo(para);
                //text += names + '\n';
            }
            return true;
//doc_builder.Write(text);
        }
    }
}
