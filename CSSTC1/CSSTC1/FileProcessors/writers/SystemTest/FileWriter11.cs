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

namespace CSSTC1.FileProcessors.writers.SystemTest {
    class FileWriter11 {
        private List<SoftwareItems> system_softwares = new List<SoftwareItems>();
        private List<DynamicHardwareItems> system_hardwares = new List<DynamicHardwareItems>();
        private Dictionary<string, List<SoftwareItems>> new_dict1 = new Dictionary<string, List<SoftwareItems>>();
        private Dictionary<string, List<DynamicHardwareItems>> new_dict2 = new
                Dictionary<string, List<DynamicHardwareItems>>();
        private List<int> time_diff = new List<int>();
        private TestEnvChartHelper test_env_helper = new TestEnvChartHelper();
        public FileWriter11(List<SoftwareItems> system_softwares, List<DynamicHardwareItems> system_hardwares) {
            this.system_softwares = system_softwares;
            this.system_hardwares = system_hardwares;
            time_diff.Add(ContentFlags.lingqucishu * 2);
            time_diff.Add(ContentFlags.pianli_1);
            time_diff.Add(ContentFlags.pianli_2);
            time_diff.Add(ContentFlags.wendangshencha);
            time_diff.Add(ContentFlags.jingtaifenxi);
            time_diff.Add(ContentFlags.daimashencha);
            time_diff.Add(ContentFlags.daimazoucha);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.pianli_3);
            time_diff.Add(ContentFlags.ceshijiuxu2);
            time_diff.Add(ContentFlags.beiceruanjianshuliang1 * 2);
            time_diff.Add(ContentFlags.beiceruanjianshuliang * 2);
            time_diff.Add(ContentFlags.peizhiceshi);
           // time_diff.Add(ContentFlags.luojiceshi);

            this.write_charts();
        }

        public void write_charts(){
            Document doc = new Document(FileConstants.save_root_file);
            DocumentBuilder doc_builder = new DocumentBuilder(doc);
            if(doc_builder.MoveToBookmark("系统测试环境标识")) {
                if(ContentFlags.test_env_id + 1 < 10)
                    doc_builder.Write( "0" + (ContentFlags.test_env_id + 1).ToString());
                else
                    doc_builder.Write((ContentFlags.test_env_id + 1).ToString());

                ContentFlags.test_env_id += 1;
            }
            this.write_bcjqd_chart(doc, doc_builder, system_softwares, system_hardwares);
            test_env_helper.write_bcjdbd_chart(doc, doc_builder, new_dict1, new_dict2, InsertionPos.xtcs_section,
                InsertionPos.xtcs_bcjdbd_sec_table, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
            test_env_helper.write_csgjhsb_chart(doc, doc_builder, new_dict1, new_dict2, 
                InsertionPos.xtcs_cshjqr_section1, InsertionPos.xtcs_csgjqr_sec_table, this.time_diff);
            test_env_helper.write_csgjhsbhcd_chart(doc, doc_builder, new_dict1, new_dict2,
                InsertionPos.xtcs_cshjhc_section1, InsertionPos.xtcs_csgjhc_sec_table, 1, this.time_diff, false);
            if(ContentFlags.xitonghuiguiceshi > 0) {
                test_env_helper.write_bcjqd_chart(doc, doc_builder, new_dict1, new_dict2, 
                    InsertionPos.xtcs_section1, InsertionPos.xtcs_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, 
                    InsertionPos.sj_bcjqd_iden_row, this.time_diff);
                test_env_helper.write_bcjdbd_chart(doc, doc_builder, new_dict1, new_dict2, 
                    InsertionPos.xtcs_section1, InsertionPos.xtcs_bcjdbd_sec_table, 
                    InsertionPos.sj_bcjdbd_name_row, InsertionPos.sj_bcjdbd_iden_row, this.time_diff);
                test_env_helper.write_csgjhsb_chart(doc, doc_builder, new_dict1, new_dict2,
                InsertionPos.xtcs_cshjqr_section2, InsertionPos.xtcs_csgjqr_sec_table, this.time_diff);
                test_env_helper.write_csgjhsbhcd_chart(doc, doc_builder, new_dict1, new_dict2,
                    InsertionPos.xtcs_cshjhc_section2, InsertionPos.xtcs_csgjhc_sec_table, 1, this.time_diff, 
                    false);
            }
            doc.Save(FileConstants.save_root_file);

        }

        //填写被测件清单
        public void write_bcjqd_chart(Document doc, DocumentBuilder doc_builder,
            List<SoftwareItems> software_list, List<DynamicHardwareItems> hardware_list) {
            //Dictionary<string, List<SoftwareItems>> new_dict1 = new Dictionary<string, List<SoftwareItems>>();
            this.new_dict1.Add(software_list[0].rj_mingcheng, software_list);      
            this.new_dict2.Add(hardware_list[0].yj_mingcheng, hardware_list);      
            test_env_helper.write_bcjqd_chart(doc, doc_builder, new_dict1, new_dict2, InsertionPos.xtcs_section,
                InsertionPos.xtcs_bcjqd_sec_table, InsertionPos.sj_bcjqd_name_row, InsertionPos.sj_bcjqd_iden_row,
                this.time_diff);
            
        }

    }
}
