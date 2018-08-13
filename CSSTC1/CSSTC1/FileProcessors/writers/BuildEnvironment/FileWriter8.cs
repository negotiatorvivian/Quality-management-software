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
    class FileWriter8 {
        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict;
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict;
        List<List<string>> files_ids = new List<List<string>>();
        List<SoftwareItems> softwares = new List<SoftwareItems>();
        public string software_names = "";
        List<int> time_diff = new List<int>();
        private TestEnvChartHelper helper;
        public FileWriter8(Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, Dictionary<string,
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
            helper = new TestEnvChartHelper();
            this.write_charts();
        }

        private void write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            this.write_lxwtf_chart_1(doc, doc_builder, ruanjianpeizhi_dict, "配置项被测件名称");
            helper.write_bcjqd_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict, 
                InsertionPos.djhj_bcjqd_section, InsertionPos.djhj_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row,
                InsertionPos.sj_bcjqd_iden_row, this.time_diff);
            helper.write_bcjdbd_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict,
                InsertionPos.djhj_bcjqd_section, InsertionPos.djhj_bcjdbd_sec_table, 
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            helper.write_bcjlqqd_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict,
                InsertionPos.djhj_bcjqd_section, InsertionPos.djhj_bcjlqqd_sec_table,
                InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, 
                InsertionPos.djhj_bcjlqqd_type_row, this.time_diff);
            helper.write_csgjhsb_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict, 
                InsertionPos.djhj_cshjqr_section, InsertionPos.djhj_cshjqr_sec_table, this.time_diff);
            ContentFlags.beiceruanjianshuliang = ruanjianpeizhi_dict.Count;
            helper.write_csgjhsbhcd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.djhj_cshjhc_section, InsertionPos.djhj_cshjhc_sec_table, 1, this.time_diff, false);
            this.time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            int num = InsertionPos.djhj_hyqdb_section;
            foreach(int i in time_diff)
                num += i;
            OperationHelper.conference_signing(doc, doc_builder, num, InsertionPos.djhj_hyqdb_sec_table);
            if(ContentFlags.pianli_3 > 0){
                this.write_wtbg_chart(doc, doc_builder, ContentFlags.jxwt_reports, num + 1, 
                    InsertionPos.djhj_pcywt_sec_table);
                this.write_wtzz_chart(doc, doc_builder, ContentFlags.jxwt_reports, num + 1, 
                    InsertionPos.djhj_pcywtzz_sec_table, InsertionPos.csdgws_solu_row);
                this.time_diff.Add(ContentFlags.pianli_3);
            }
                
            doc.Save(FileConstants.save_root_file);
        }

        public void write_lxwtf_chart_1(Document doc, DocumentBuilder doc_builder,
            Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict, string bookmark) {
            List<string> softwares = ruanjianpeizhi_dict.Select(r => r.Key).ToList();
            string text = "";
            foreach(string name in softwares){
                text += name + '、';
            }
            text = text.Substring(0, text.Length - 1);
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(text);
        }

        //项目SQA偏差与问题报告
        public bool write_wtbg_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files,
            int section_index, int sec_table_index) {
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, InsertionPos.csdgws_ques_row, 0);
                doc_builder.Write(report.wenti);
                index_row += 1;
            }
            return true;
        }

        //偏差与问题追踪
        public bool write_wtzz_chart(Document doc, DocumentBuilder doc_builder, List<QestionReport> files,
            int section_index, int sec_table_index, int row_index) {
            doc_builder.MoveToSection(section_index);
            Table table = (Table)doc.Sections[section_index].GetChild(NodeType.Table,
                sec_table_index, true);
            int index_row = 3;
            foreach(QestionReport report in files) {
                if(index_row < files.Count + 2) {
                    var row = table.Rows[index_row].Clone(true);
                    table.Rows.Insert(index_row + 1, row);
                }
                doc_builder.MoveToCell(sec_table_index, index_row, row_index, 0);
                doc_builder.Write(report.chulicuoshi);
                index_row += 1;
            }
            return true;
        }
    }
}
