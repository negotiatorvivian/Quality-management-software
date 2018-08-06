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
        public static bool peizhiceshi;
        public static bool xitongceshi;
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
        public static Dictionary<string, SoftwareItems> software_dict;
        //第一部分上传表格中选择静态分析的软件名称集合
        public static List<string> software_list = new List<string>();
        //软件静态测试环境软件项
        public static List<StaticAnalysisFile> static_files = new List<StaticAnalysisFile>();
        //被测件子软件数量
        public static int beiceruanjianshuliang = 1;
        //是否有测试就绪内部评审
        public static int pianli_3 = 1;
        #endregion

    }
}
