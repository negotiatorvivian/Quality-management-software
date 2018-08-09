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
            helper = new TestEnvChartHelper(ruanjianpeizhi_dict, yingjianpeizhi_dict);
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
            doc.Save(FileConstants.save_root_file);
            return true;
        }
    }
}
