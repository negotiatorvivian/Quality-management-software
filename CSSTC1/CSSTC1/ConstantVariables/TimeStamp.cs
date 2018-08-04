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

        //第三阶段偏离联系时间
        public static string pianli2_time;

        //就绪内部评审时间
        public static string csjxps_time;
        public static string csjxps_format_time;

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
        

        public static List<TestEnvironment> test_envs = new List<TestEnvironment>();

    }
}
