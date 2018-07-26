using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CSSTC1.ConstantVariables {
    class NamingRules {
        public static string Xm_biaoshi = "项目标识";
        public static string Xm_mingcheng = "项目名称";
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

        public static string ignore_file = "中国船舶工业集团公司软件质量与可靠性测评中心";
        public static string company = "中国船舶工业集团公司软件质量与可靠性测评中心";

        public static string[] PzControls = { "文档审查", "静态分析", "功能测试", "性能测试", "强度测试", "接口测试", "可靠性测                      试", "边界测试", "安全性测试", "人机交互界面测试", "代码走查", "容量测试", "余量测试", "安装性测试", "敏感                       性测试", "恢复性测试", "互操作性测试", "兼容性测试", "标准符合性测试", "数据处理测试", "内存使用缺陷", "逻                      辑测试" };

        public static string[] Xmjj_paras = { "xt_mingcheng", "rj_mingcheng", "xt_banben", "yx_huanjing", "kf_huanjing", "bc_yuyan", "dm_guimo", "dengji", "yz_danwei", "beizhu" };

        public static string[] kfhj_params = { "Visual", "Tornado" };
        public static string[] kfhj_bookmarks = { "VC", "tornado", "BC", "编程环境其他" };

        public static Dictionary<int, string> table_names = new Dictionary<int, string>();
        public static Dictionary<string, string> xmjj_names = new Dictionary<string, string>();

        public static string pre_name = "CSSTC-TW-";

        public Dictionary<string, int> file_type_match = new Dictionary<string, int>();

        public static string get_file_type(string name) {
            Regex reg = new Regex(@"[\u4e00-\u9fa5]{4,}");
            Match match = reg.Match(name, 0);
            string title = match.Groups[0].Value;
            switch(title) {
                case "软件研制总要求":
                    return "01";
                case "软件研制任务书":
                    return "02";
                case "软件任务书":
                    return "02";
                case "软件质量保证计划":
                    return "03";
                case "软件配置管理计划":
                    return "04";
                case "软件需求规格说明":
                    return "05";
                case "接口需求规格说明":
                    return "06";
                case "软件设计文档":
                    return "07";
                case "软件概要设计文档":
                    return "08";
                case "软件概要设计报告":
                    return "08";
                case "软件详细设计文档":
                    return "09";
                case "软件详细设计说明书":
                    return "09";
                case "详细设计说明书":
                    return "09";
                case "接口设计文档":
                    return "10";
                case "软件产品规格说明":
                    return "11";
                case "版本说明文档":
                    return "12";
                case "软件测试计划":
                    return "13";
                case "软件测试说明":
                    return "14";
                case "软件测试报告":
                    return "15";
                case "计算机系统操作员手册":
                    return "16";
                case "软件用户手册":
                    return "17";
                case "软件程序员手册":
                    return "18";
                case "软件代码":
                    return "19";
                case "被测系统硬件":
                    return "20";
                case "测试协议":
                    return "21";
                case "测试合同":
                    return "22";
                case "被测件清单":
                    return "23";
                case "被测件光盘":
                    return "24";
                case "软件安装包":
                    return "26";
                default:
                    return "25";
            }

        }

        public static string get_job_title(string name) {
            switch(name) {
                case "陈大圣":
                    return "主任";
                case "唐龙利":
                    return "常务副主任";
                case "韩新宇":
                    return "副主任";
                case "张凯":
                    return "高工";
                default:
                    return "工程师";
            }
        }
    }
}
