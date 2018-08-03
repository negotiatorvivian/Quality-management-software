using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CSSTC1.ConstantVariables {
    class NamingRules {
        public static string Xm_biaoshi = "项目标识";
        public static string Xm_mingcheng = "项目名称";
        public static string Rj_mingcheng = "软件名称";
        public static string Pz_guanliyuan = "配置管理员";
        public static string Bcj_guanliyuan = "被测件管理员";
        public static string Zlbz_renyuan = "质量保证人员";
        public static string Cs_jianduyuan = "测试监督员";
        public static string Sb_guanliyuan = "设备管理员";
        public static string Fuzhurren = "副主任";
        public static string Cw_fuzhuren = "常务副主任";
        public static string Zhuren = "主任";
        public static string Wt_xingming = "委托方姓名";
        public static string Wt_dianhua = "委托方电话";
        public static string Wt_danwei = "委托方单位";
        public static string Kf_xingming = "开发方姓名";
        public static string Kf_dianhua = "开发方电话";
        public static string Kf_danwei = "系统（软件）开发单位";
        public static string Xm_kaishishijian = "项目开始时间";
        public static string Xm_jieshushijian = "项目结束时间";

        public static string[] Csyj_files = {"研制任务书", "技术方案", "技术规格书", "软件研制任务书",
                                                "软件需求规格说明"};

        public static string ignore_file = "中国船舶工业集团公司软件质量与可靠性测评中心";
        public static string company = "中国船舶工业集团公司软件质量与可靠性测评中心";

        public static string[] PzControls = { "文档审查", "静态分析", "功能测试", "性能测试", "强度测试", "接口测试", 
                                                "可靠性测试", "边界测试", "安全性测试", "人机交互界面测试", "代码走查",
                                                "容量测试", "余量测试", "安装性测试", "敏感性测试", "恢复性测试", 
                                                "互操作性测试", "兼容性测试", "标准符合性测试", "数据处理测试", 
                                                "内存使用缺陷", "逻辑测试" };

        public static string[] Xmjj_paras = { "xt_mingcheng", "rj_mingcheng", "xt_banben", "yx_huanjing", 
                                                "kf_huanjing", "bc_yuyan", "dm_guimo", "dengji", "yz_danwei", 
                                                "beizhu" };

        public static string[] kfhj_params = { "Visual", "Tornado" };
        public static string[] kfhj_bookmarks = { "VC", "tornado", "BC", "编程环境其他" };

        public static Dictionary<int, string> table_names = new Dictionary<int, string>();
        public static Dictionary<string, string> xmjj_names = new Dictionary<string, string>();

        public static string pre_name = "CSSTC-TW-";

        public Dictionary<string, int> file_type_match = new Dictionary<string, int>();

        public static string pingshenzuzhang = "唐龙利";

        
    }
}
