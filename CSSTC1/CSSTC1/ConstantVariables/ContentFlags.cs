using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;

namespace CSSTC1.ConstantVariables {
    class ContentFlags {
        //缺失值指示变量
        public static int missing = -1;
        //评审组成员
        public static List<string> pingshenzuchengyuan = new List<string>();
        //
        public static int rk_id = 1;
        public static int ck_id = 1;
        public static int dbd_id = 1;
        public static int lxwtf_id = 2;
        //联系委托方时间
        public static Dictionary<string, DateTime> time_dict1 = new Dictionary<string, DateTime>();
        //入库等时间
        public static Dictionary<string, DateTime> time_dict2 = new Dictionary<string, DateTime>();
        #region 项目立项阶段
        public static float code_line = 0;
        //配置项测试节数
        public static int peizhiceshi = 12;
        //系统首轮测试
        public static int xitongceshi = 10;
        public static List<ProjectInfo> pro_infos = new List<ProjectInfo>();
        //所有的文档（包含多个版本）
        public static List<FileList> all_file_lists = new List<FileList>();
        //测试类型
        public static HashSet<string> test_types = new HashSet<string>();
        //文档命名计数
        public static Dictionary<string, int> file_counter = new Dictionary<string, int>();
        public static Dictionary<string, int> file_type_counter = new Dictionary<string, int>();
        //立项阶段被测件清单
        public static Dictionary<string, FileList> beicejianqingdan_dict = new Dictionary<string, FileList>();
        //若领取两次，每一次领取的数量
        public static List<int> beicejianshuliang = new List<int>();
        //入库文档合集（key为文档名称）
        public static Dictionary<string, string> rukuwendang_dict = new Dictionary<string, string>();
        //被测件领取次数
        public static int lingqucishu = 2;

        //是否偏离
        public static int pianli_1 = 1;
        #endregion

        #region 测试需求分析及测试策划阶段
        //是否选择测试大纲
        public static bool ceshidagang = true;
        //是否偏离
        public static int pianli_2 = 1;

        #endregion

        #region 测试设计和执行阶段

        public static Dictionary<string, StaticAnalysisFile> software_dict;
        //第一部分上传表格中选择静态分析的软件名称集合
        public static List<string> static_list = new List<string>();
        public static List<StaticAnalysisFile> static_files = new List<StaticAnalysisFile>();
        //第一部分上传表格中选择代码走查的软件名称集合
        public static List<string> dmzc_software_list = new List<string>();
        public static List<StaticAnalysisFile> dmzc_software_info = new List<StaticAnalysisFile>();
        //第一部分上传表格中选择代码审查的软件名称集合
        public static List<string> dmsc_software_list = new List<string>();
        public static List<StaticAnalysisFile> dmsc_software_info = new List<StaticAnalysisFile>();
        //软件动态测试环境软件项
        public static List<string> dynamic_list = new List<string>();
        public static List<StaticAnalysisFile> dynamic_files = new List<StaticAnalysisFile>();
        //上传文件中的测试用例表格
        public static Dictionary<string, List<TestExample>> example_dict = new Dictionary
            <string, List<TestExample>>();
        //上传文件中的测试问题统计表格
        public static Dictionary<string, List<TestProblem>> problems = new Dictionary<string, List<TestProblem>>();
        //上传文件中的系统测试问题统计表格
        public static Dictionary<string, List<TestProblem>> system_problems = new Dictionary<string, 
            List<TestProblem>>();
        //被测件子软件数量
        public static int beiceruanjianshuliang = 0;
        //第二次就绪子软件数量
        public static int beiceruanjianshuliang1 = 0;
        //就绪存在问题意见表
        public static List<QestionReport> jxwt_reports = new List<QestionReport>();
        //软件测试环境
        public static List<TestEnvironment> test_envs = new List<TestEnvironment>();
        //软件的最后一次测试环境（配置测试用）
        public static Dictionary<string, List<SoftwareItems>> ruanjianpeizhi_dict = new Dictionary<string,
                List<SoftwareItems>>();
        public static Dictionary<string, List<DynamicHardwareItems>> yingjianpeizhi_dict = new Dictionary<string,
            List<DynamicHardwareItems>>();
        //系统测试软件环境与硬件环境
        public static List<SoftwareItems> system_softwares = new List<SoftwareItems>();
        public static List<DynamicHardwareItems> system_hardwares = new List<DynamicHardwareItems>();
        public static int test_env_id = 0;
        public static Dictionary<string, string> test_env_id_dict = new Dictionary<string, string>();

        public static Dictionary<string, string> sys_time_dict = new Dictionary<string,string>();
        //是否有偏离报告
        public static int pianli_3 = 1;
        //是否有文档审查
        public static int wendangshencha = 6;
        //是否有静态分析
        public static int jingtaifenxi = 15;
        //是否有代码审查(13节)
        public static int daimashencha = 13;
        //代码审查是否与静态分析范围一致
        public static bool dmsc_same = false;
        //是否有代码走查(13节)
        public static int daimazoucha = 13;
        //代码走查是否与静态分析范围一致
        public static bool dmzc_same = false;
        //是否有第二次测试就绪
        public static int ceshijiuxu2 = 3;
        //是否有逻辑测试
        public static int luojiceshi = 3;
        //是否有系统回归测试
        public static int xitonghuiguiceshi = 10;

        //测试与执行阶段是否有偏离
        public static int pianli_4 = 1;
        #endregion

        #region 测试总结阶段
        #endregion

        public static Dictionary<string, List<string>> set_order() {
            //Dictionary<string, List<string>> lxwtf_time_dict = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> project_id_dict = new Dictionary<string,List<string>>();
            List<List<string>> date_bookmarks = new List<List<string>>();
            for(int j = 0; j < 40; j ++){
                date_bookmarks.Add(new List<string>());
            }
            date_bookmarks[0].Add("调拨1");
            date_bookmarks[0].Add("入库1");
            date_bookmarks[0].Add("配置41");
            project_id_dict.Add("入库日期1", date_bookmarks[0]);
            date_bookmarks[1].Add("调拨2");
            date_bookmarks[1].Add("入库2");
            date_bookmarks[1].Add("配置42");
            project_id_dict.Add("入库日期2", date_bookmarks[1]);
            date_bookmarks[2].Add("入库3");
            date_bookmarks[2].Add("配置40");
            project_id_dict.Add("内部评审时间", date_bookmarks[2]);
            date_bookmarks[3].Add("出库1");
            date_bookmarks[3].Add("入库4");
            date_bookmarks[3].Add("配置39");
            project_id_dict.Add("测试大纲出库时间", date_bookmarks[3]);
            date_bookmarks[4].Add("入库5");
            date_bookmarks[4].Add("配置38");
            project_id_dict.Add("测试说明与计划内审时间", date_bookmarks[4]);
            date_bookmarks[5].Add("入库6");
            date_bookmarks[5].Add("出库2");
            date_bookmarks[5].Add("配置37");
            project_id_dict.Add("测试说明与计划出库时间", date_bookmarks[5]);
            date_bookmarks[6].Add("入库7");
            date_bookmarks[6].Add("配置36");
            project_id_dict.Add("文档审查结果记录入库时间", date_bookmarks[6]);
            date_bookmarks[7].Add("调拨3");
            date_bookmarks[7].Add("入库8");
            date_bookmarks[7].Add("出库3");
            date_bookmarks[7].Add("配置35");
            project_id_dict.Add("文档审查回归时间", date_bookmarks[7]);
            date_bookmarks[8].Add("入库9");
            date_bookmarks[8].Add("调拨4");
            date_bookmarks[8].Add("配置34");
            date_bookmarks[8].Add("配置1");
            project_id_dict.Add("静态分析审查时间", date_bookmarks[8]);
            date_bookmarks[9].Add("入库10");
            date_bookmarks[9].Add("配置2");
            project_id_dict.Add("静态分析结果入库时间", date_bookmarks[9]);
            date_bookmarks[10].Add("调拨5");
            date_bookmarks[10].Add("出库4");
            date_bookmarks[10].Add("入库11");
            date_bookmarks[10].Add("配置3");
            date_bookmarks[10].Add("入库12");
            date_bookmarks[10].Add("配置4");

            project_id_dict.Add("静态分析回归时间", date_bookmarks[10]);
            int i = project_id_dict.Count;
            if(ContentFlags.daimashencha > 0) {
                if(!ContentFlags.dmsc_same){
                    date_bookmarks[i].Add("调拨6");
                    //date_bookmarks[i].Add("配置5");
                    date_bookmarks[i].Add("配置5_");
                    project_id_dict.Add("代码审查时间", date_bookmarks[i]);
                    i += 1;
                    date_bookmarks[i].Add("调拨7");
                    date_bookmarks[i].Add("入库13");
                    date_bookmarks[i].Add("入库15");
                    //date_bookmarks[12].Add("出库5");
                    project_id_dict.Add("代码审查回归时间", date_bookmarks[i]);
                    i += 1;
                }

                date_bookmarks[i].Add("入库14");
                //date_bookmarks[i].Add("出库5");
                date_bookmarks[i].Add("配置6");
                project_id_dict.Add("代码审查文档入库时间", date_bookmarks[i]);
                i += 1;
                date_bookmarks[i].Add("配置5");
                project_id_dict.Add("代码审查时间", date_bookmarks[i]);
                i += 1;
                date_bookmarks[i].Add("入库16");
                date_bookmarks[i].Add("出库5");
                date_bookmarks[i].Add("配置7");
                date_bookmarks[i].Add("配置8");
                project_id_dict.Add("代码审查回归时间", date_bookmarks[i]);
                i += 1;
            }
            if(ContentFlags.daimazoucha > 0) {
                if(!ContentFlags.dmzc_same) {
                    date_bookmarks[i].Add("调拨8");
                    date_bookmarks[i].Add("入库17");
                    date_bookmarks[i].Add("配置9");
                    project_id_dict.Add("代码走查时间", date_bookmarks[i]);
                    i += 1;
                    date_bookmarks[i].Add("调拨9");
                    date_bookmarks[i].Add("入库19");
                    project_id_dict.Add("代码走查回归时间", date_bookmarks[i]);
                    i += 1;
                }
                date_bookmarks[i].Add("配置10");
                project_id_dict.Add("代码走查时间", date_bookmarks[i]);
                i += 1;
                
                date_bookmarks[i].Add("出库6");
                date_bookmarks[i].Add("入库20");
                date_bookmarks[i].Add("配置12");
                date_bookmarks[i].Add("配置13");
                project_id_dict.Add("代码走查回归时间", date_bookmarks[i]);
                i += 1;
                date_bookmarks[i].Add("入库18");
                date_bookmarks[i].Add("配置11");
                project_id_dict.Add("代码走查入库申请时间", date_bookmarks[i]);
                i += 1;
            }
            date_bookmarks[i].Add("入库21");
            date_bookmarks[i].Add("配置14");
            project_id_dict.Add("测试说明内部评审时间", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("入库22");
            date_bookmarks[i].Add("调拨10");
            date_bookmarks[i].Add("配置15");
            date_bookmarks[i].Add("配置16");
            project_id_dict.Add("搭建环境被测件接收时间", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("入库23");
            date_bookmarks[i].Add("调拨11");
            date_bookmarks[i].Add("配置17");
            date_bookmarks[i].Add("配置18");
            project_id_dict.Add("搭建环境被测件接收时间1", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("入库25");
            date_bookmarks[i].Add("调拨12");
            date_bookmarks[i].Add("配置20");
            date_bookmarks[i].Add("配置21");
            project_id_dict.Add("配置项动态回归被测件接收时间", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("配置19");
            date_bookmarks[i].Add("入库24");
            project_id_dict.Add("联系委托方第八次", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("入库26");
            date_bookmarks[i].Add("入库26引用");
            date_bookmarks[i].Add("配置22");
            date_bookmarks[i].Add("出库7");
            date_bookmarks[i].Add("配置23");
            project_id_dict.Add("回归测试时间", date_bookmarks[i]);
            i += 1;
            if(xitongceshi > 0){
                date_bookmarks[i].Add("入库27");
                date_bookmarks[i].Add("入库27引用1");
                date_bookmarks[i].Add("出库8");
                date_bookmarks[i].Add("调拨13");
                date_bookmarks[i].Add("配置24");
                date_bookmarks[i].Add("配置25");
                date_bookmarks[i].Add("配置26");
                project_id_dict.Add("系统测试时间", date_bookmarks[i]);
                i += 1;
            }
            if(xitonghuiguiceshi > 0){
                date_bookmarks[i].Add("入库28");
                date_bookmarks[i].Add("入库32");
                date_bookmarks[i].Add("配置27");
                date_bookmarks[i].Add("配置28");
                date_bookmarks[i].Add("配置29");
                date_bookmarks[i].Add("调拨14");
                project_id_dict.Add("系统回归测试时间", date_bookmarks[i]);
                i += 1;
                date_bookmarks[i].Add("出库9");
                project_id_dict.Add("系统回归测试结束时间", date_bookmarks[i]);
                i += 1;
            }
            date_bookmarks[i].Add("入库29");
            date_bookmarks[i].Add("配置30");
            project_id_dict.Add("鉴定测评报告时间", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("入库30");
            date_bookmarks[i].Add("出库10");
            date_bookmarks[i].Add("配置31");
            project_id_dict.Add("鉴定测评报告外审时间", date_bookmarks[i]);
            i += 1;
            date_bookmarks[i].Add("出库11");
            date_bookmarks[i].Add("入库31");
            date_bookmarks[i].Add("出库12");
            date_bookmarks[i].Add("配置32");
            date_bookmarks[i].Add("配置33");
            project_id_dict.Add("鉴定测评报告出库时间", date_bookmarks[i]);
            i += 1;
            return project_id_dict;
        }

        public static Dictionary<string, string> set_lxwtf_order() {
            Dictionary<string, string> lxwtf_time_dict = new Dictionary<string,string>();
            lxwtf_time_dict.Add("偏离联系时间", "联系委托方3");
            lxwtf_time_dict.Add("联系委托方第三次", "联系委托方4");
            lxwtf_time_dict.Add("联系委托方第四次", "联系委托方5");
            lxwtf_time_dict.Add("偏离联系时间2", "联系委托方6");
            lxwtf_time_dict.Add("文档审查确认时间", "联系委托方7");
            lxwtf_time_dict.Add("联系委托方第五次", "联系委托方8");
            lxwtf_time_dict.Add("联系委托方第六次", "联系委托方9");
            lxwtf_time_dict.Add("代码审查联系委托方", "联系委托方10");
            lxwtf_time_dict.Add("代码审查文档入库时间", "联系委托方11");
            lxwtf_time_dict.Add("代码走查联系委托方", "联系委托方12");
            lxwtf_time_dict.Add("代码走查入库申请时间", "联系委托方13");
            lxwtf_time_dict.Add("联系委托方第七次", "联系委托方14");
            lxwtf_time_dict.Add("联系委托方第八次", "联系委托方15");
            lxwtf_time_dict.Add("联系委托方第九次", "联系委托方16");
            lxwtf_time_dict.Add("联系委托方第十次", "联系委托方17");
            lxwtf_time_dict.Add("联系委托方第十一次", "联系委托方18");
            lxwtf_time_dict.Add("联系委托方第十二次", "联系委托方19");
            lxwtf_time_dict.Add("联系委托方第十三次", "联系委托方20");
            lxwtf_time_dict.Add("偏离联系时间3", "联系委托方21");
            return lxwtf_time_dict;
        }
    }
}
