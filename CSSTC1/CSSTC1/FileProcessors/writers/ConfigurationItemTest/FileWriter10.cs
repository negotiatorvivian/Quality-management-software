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
using CSSTC1.FileProcessors.writers.BuildEnvironment;

namespace CSSTC1.FileProcessors.writers.ConfigurationItemTest {
    class FileWriter10 {
        private List<int> time_diff = new List<int>();
        private TestEnvChartHelper helper = new TestEnvChartHelper();

        Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict = new Dictionary<string,
                List<SoftwareItems>>();
        Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict = new Dictionary<string,
            List<DynamicHardwareItems>>();
        private ConfigChartHelper config_helper = new ConfigChartHelper();
        public FileWriter10(Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict,
            Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict) {
            time_diff.Add(ContentFlags.lingqucishu * 2);
            time_diff.Add(ContentFlags.pianli_1);
            time_diff.Add(ContentFlags.pianli_2);
            time_diff.Add(ContentFlags.wendangshencha);
            time_diff.Add(ContentFlags.jingtaifenxi);
            time_diff.Add(ContentFlags.daimashencha);
            time_diff.Add(ContentFlags.daimazoucha);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.pianli_3);
            time_diff.Add(ContentFlags.beiceruanjianshuliang1 * 2);
            time_diff.Add(ContentFlags.ceshijiuxu2);
            this.ruanjianpeizhi_dict = ruanjianpeizhi_dict;
            this.yingjianpeizhi_dict = yingjianpeizhi_dict;
            this.write_charts();
        }

        public void write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            config_helper.write_rksqd_chart(doc, doc_builder, InsertionPos.pzcs_rksqd_section, 
                InsertionPos.pzcs_rksqd_sec_table,InsertionPos.sj_rksqd_name_row, this.time_diff);
            Dictionary<string, FileList> dict = ConfigChartHelper.set_file_id(doc,doc_builder, 
                ContentFlags.pro_infos);
            ChartHelper.write_bcjqd_chart(doc, doc_builder, dict, InsertionPos.pzcs_bcjqd_section,
                InsertionPos.pzcs_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row,
                this.time_diff);
            ChartHelper.write_bcjdbd_chart(doc, doc_builder, dict, InsertionPos.pzcs_bcjqd_section,
                InsertionPos.pzcs_bcjdbd_sec_table, 1, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            ChartHelper.write_bcjlqqd_chart(doc, doc_builder, dict, InsertionPos.pzcs_bcjqd_section,
                InsertionPos.pzcs_bcjlqqd_sec_table, 2, InsertionPos.sj_bcjdbd_name_row,
                InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            //测试环境确认与测试环境核查单
            this.write_cshj_chart(doc, doc_builder);
            config_helper.write_rksqd_chart1(doc, doc_builder, InsertionPos.pzcs_rksqd_section1,
                InsertionPos.pzcs_rksqd_sec_table, InsertionPos.sj_rksqd_name_row, this.time_diff);
            config_helper.write_rksqd_chart2(doc, doc_builder, InsertionPos.pzcs_rksqd_section2,
                InsertionPos.pzcs_rksqd_sec_table, InsertionPos.sj_rksqd_name_row, this.time_diff);
            doc.Save(FileConstants.save_root_file);
        }

        public void write_cshj_chart(Document doc, DocumentBuilder doc_builder){
            if(TimeStamp.csjxps_format_time.Count == 1){
                helper.write_csgjhsb_chart(doc, doc_builder, this.ruanjianpeizhi_dict, this.yingjianpeizhi_dict,
                    InsertionPos.pzcs_cshjqr_section1, InsertionPos.djhj_cshjqr_sec_table, this.time_diff);
                helper.write_csgjhsbhcd_chart1(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.pzcs_cshjhc_section1, InsertionPos.djhj_cshjhc_sec_table, 1, this.time_diff);
            }
            else{
                Dictionary<string, List<SoftwareItems>> softwares = new Dictionary<string,List<SoftwareItems>>();
                Dictionary<string, List<DynamicHardwareItems>> hardwares = new 
                    Dictionary<string,List<DynamicHardwareItems>>();
                foreach(string key in ContentFlags.ruanjianpeizhi_dict.Keys){
                    if(this.ruanjianpeizhi_dict.ContainsKey(key)){
                        softwares[key] = this.ruanjianpeizhi_dict[key];
                        hardwares[key] = this.yingjianpeizhi_dict[key];
                    }
                    else{
                        softwares.Add(key, ContentFlags.ruanjianpeizhi_dict[key]);
                        hardwares.Add(key, ContentFlags.yingjianpeizhi_dict[key]);
                    }
                }
                helper.write_csgjhsb_chart(doc, doc_builder, softwares, hardwares, 
                    InsertionPos.pzcs_cshjqr_section1, InsertionPos.djhj_cshjqr_sec_table, this.time_diff);
                helper.write_csgjhsbhcd_chart1(doc, doc_builder, softwares, hardwares, 
                    InsertionPos.pzcs_cshjhc_section1, InsertionPos.djhj_cshjhc_sec_table, 1, this.time_diff);
            }
            this.time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
        }

        
    }
}
