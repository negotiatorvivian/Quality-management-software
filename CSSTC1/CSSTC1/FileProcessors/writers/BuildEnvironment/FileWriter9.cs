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
        Dictionary<string, List<string>> software_test_env2 = new Dictionary<string, List<string>>();
        List<List<string>> files_ids = new List<List<string>>();
        List<SoftwareItems> softwares = new List<SoftwareItems>();
        public string software_names = "";
        List<int> time_diff = new List<int>();
        private TestEnvChartHelper helper;
        private string project_id;
        private string year;
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
            this.project_id = doc.Range.Bookmarks["项目标识"].Text;
            this.year = doc.Range.Bookmarks["年份"].Text;


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
                OperationHelper.delete_table(doc, doc_builder, "搭建环境被测件领取清单", 2);
            helper.write_bcjlqqd_chart(doc, doc_builder, temp1, temp2, InsertionPos.djhj_bcjqd_section1, 
                InsertionPos.djhj_bcjlqqd_sec_table, InsertionPos.sj_bcjdbd_name_row, 
                InsertionPos.sj_bcjdbd_iden_row, InsertionPos.djhj_bcjlqqd_type_row, this.time_diff);
            this.write_pzztbbd_chart(doc, doc_builder, "测试就绪软件名称2");
            helper.write_csgjhsb_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.djhj_cshjqr_section1, InsertionPos.djhj_cshjqr_sec_table, this.time_diff);
            helper.write_csgjhsbhcd_chart(doc, doc_builder, ruanjianpeizhi_dict, yingjianpeizhi_dict,
                InsertionPos.djhj_cshjhc_section1, InsertionPos.djhj_cshjhc_sec_table, 1, this.time_diff, true);
            ContentFlags.beiceruanjianshuliang1 = ruanjianpeizhi_dict.Count;
            this.time_diff.Add(ContentFlags.beiceruanjianshuliang1 * 2);

            Dictionary<string, List<string>> software_test_env1 = this.merge_location(doc, doc_builder);
            this.write_cshj_chart(doc, doc_builder, this.software_test_env2, "搭建环境配置状态报告单1");
            this.write_cshj_chart(doc, doc_builder, software_test_env1, "配置项动态回归测试配置状态报告单2");
            int index = InsertionPos.djhj_hyqdb_section1;
            foreach(int i in time_diff)
                index += i;
            OperationHelper.conference_signing(doc, doc_builder, index, InsertionPos.djhj_hyqdb_sec_table);
            doc.Save(FileConstants.save_root_file);
            return true;
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

        public Dictionary<string, List<string>> merge_location(Document doc, DocumentBuilder doc_builder) {
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
            foreach(string key in software_test_envs.Keys)
                this.software_test_env2.Add(key, software_test_envs[key]);
            List<string> jiuxu_softwares2 = ContentFlags.test_envs.Select(r => r.rj_mingcheng).ToList();
            foreach(ProjectInfo info in ContentFlags.pro_infos){
                if(jiuxu_softwares2.Contains(info.rj_mingcheng))
                    continue;
                else{
                    if(software_test_envs.ContainsKey("本测评中心")){
                        software_test_envs["本测评中心"].Add(info.rj_mingcheng);
                    }
                    else{
                        List<string> temp = new List<string>();
                        temp.Add(info.rj_mingcheng);
                        software_test_envs.Add("本测评中心", temp);
                    }
                }
            }
            return software_test_envs;
        }

        //配置报告单中的测试环境
        public bool write_cshj_chart(Document doc, DocumentBuilder doc_builder, Dictionary<string, 
            List<string>> software_test_envs, string bookmark) {
            if(!doc_builder.MoveToBookmark(bookmark))
                return false;
            
            doc_builder.MoveToBookmark(bookmark);
            string contents = "";

            foreach(string location in software_test_envs.Keys) {
                List<string> test_env_id_list = new List<string>();
                string names = "";
                if(!location.Equals("本测评中心")){
                    foreach(string software in software_test_envs[location]) {  
                        names += software + "、";
                    }
                    test_env_id_list = this.get_test_env_id_list(location, software_test_envs[location].Count);
                    
                }
                else{
                    foreach(string software in software_test_envs[location]) {
                        names += software + "、";
                        test_env_id_list.Add(ContentFlags.test_env_id_dict[software]);
                    }
                }
                names = names.Substring(0, names.Length - 1);
                names += "动态测试环境已搭建，位于" + location + "，测试环境标识：";
                foreach(string text in test_env_id_list)
                    names += text + "、";
                names = names.Substring(0, names.Length - 1);
                names += "，测试环境所用测试设备、工具等的标识与测试说明中一致;\n";
                contents += names;
            }
            contents = contents.Substring(0, contents.Length - 1);
            if(doc_builder.MoveToBookmark(bookmark))
                doc_builder.Write(contents);
            return true;
        }

        public List<string> get_test_env_id_list(string location, int count){
            List<string> test_env_id_list = new List<string>();
            string head = "CSSTC-TE-{" + this.project_id + "}-N";
            string tail = "-" + this.year;
            if(location.Equals("本测评中心")) {
                //int start_index = ContentFlags.wendangshencha/6 + ContentFlags.jingtaifenxi/15;
                
                //for(int i = 0; i < count; i++) {
                //    start_index += 1;
                //    string text = head;
                //    if(start_index < 10)
                //        text += "0" + ContentFlags.test_env_id.ToString() + tail;
                //    else
                //        text += start_index.ToString() + tail;
                //    test_env_id_list.Add(text);
                //}
            }
            else{
                for(int i = 0; i < count; i++){
                    ContentFlags.test_env_id += 1;
                    string text = head;
                    if(ContentFlags.test_env_id < 10)
                        text += "0" + ContentFlags.test_env_id.ToString() + tail;
                    else
                        text += ContentFlags.test_env_id.ToString() + tail;
                    test_env_id_list.Add(text);
                }
            }
            return test_env_id_list;
        }
    }
}
