using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;

namespace CSSTC1.ConstantVariables {
    public class TimeStamp {
        public static string xiangmukaishishijian = "";
        public static List<string> lingqushijian = new List<string>();
        public static string ceshins_time;

        public static string ceshisc_time;
        public static string ceshisc_format_time;

        //首轮动态测试时间
        public static string sldtcs_time;
        public static string sldtcs_format_time;

        //第三阶段偏离联系时间
        public static string pianli2_time;


        //就绪内部评审时间
        public static List<string> csjxps_time = new List<string>();
        public static List<string> csjxps_format_time = new List<string>();

        //文档审查确认时间
        public static string wdscqr_time;
        public static string wdscqr_format_time;

        //文档审查回归时间
        public static string wdschg_time;
        public static string wdschg_format_time;

        //静态分析审查时间
        public static string jtfxsc_time;
        public static string jtfxsc_format_time;

        //静态分析回归时间
        public static string jtfxhg_time;
        public static string jtfxhg_format_time;
        //静态分析确认时间
        public static string jtfxqr_time;
        public static string jtfxqr_format_time;

        //代码审查时间
        public static string dmsc_time;
        public static string dmsc_format_time;
        //代码审查回归时间
        public static string dmschg_time;
        public static string dmschg_format_time;
        //代码审查确认时间
        public static string dmscqr_time;
        public static string dmscqr_format_time;

        //代码走查时间
        public static string dmzc_time;
        public static string dmzc_format_time;
        //代码走查回归时间
        public static string dmzchg_time;
        public static string dmzchg_format_time;
        //代码走查确认时间
        public static string dmzcqr_time;
        public static string dmzcqr_format_time;

        //逻辑测试时间
        public static string ljcs_time;
        public static string ljcs_format_time;

        //首轮系统测试时间
        public static string slxtcs_time;
        public static string slxtcs_format_time;
        //系统回归测试时间
        public static string xthgcs_time;
        public static string xthgcs_format_time;

        //测试报告评审时间
        public static string csbgps_time;
        public static string csbgps_format_time;
       //测试说明评审时间
        public static string cssmps_time;
        public static string cssmps_format_time;
        //就绪内部评审时间
        public static string jxnbps_time;
        public static string jxnbps_format_time;

    }
}
