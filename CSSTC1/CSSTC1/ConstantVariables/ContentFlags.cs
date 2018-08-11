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
        #region 项目立项阶段
        public static int peizhiceshi = 12;
        public static int xitongceshi = 20;
        public static List<ProjectInfo> pro_infos = new List<ProjectInfo>();
        public static List<FileList> all_file_lists = new List<FileList>();
       

        //测试类型
        public static HashSet<string> test_types = new HashSet<string>();
        //立项阶段被测件清单
        public static Dictionary<string, FileList> beicejianqingdan_dict = new Dictionary<string, FileList>();
        //被测件领取次数
        public static int lingqucishu = 2;
        ////指示节复制次数
        //public static int copy = 0;
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

        //是否有偏离报告
        public static int pianli_3 = 1;
        //是否有文档审查
        public static int wendangshencha = 6;
        //是否有静态分析
        public static int jingtaifenxi = 15;
        //是否有代码审查(13节)
        public static int daimashencha = 13;
        //是否有代码走查(13节)
        public static int daimazoucha = 13;
        //是否有第二次测试就绪
        public static int ceshijiuxu2 = 3;
        //是否有逻辑测试
        public static int luojiceshi = 3;
        //是否有系统回归测试
        public static int xitonghuiguiceshi = 10;
        #endregion

        #region 测试总结阶段
        #endregion
    }
}
