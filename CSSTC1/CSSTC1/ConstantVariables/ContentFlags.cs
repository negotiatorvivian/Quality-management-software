using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSSTC1.FileProcessors.models;

namespace CSSTC1.ConstantVariables {
    class ContentFlags {
        public static bool peizhiceshi;
        public static bool xitongceshi;
        public static int missing = -1;
        public static int copy = 0;

        public static HashSet<string> test_types = new HashSet<string>();
        public static List<StaticAnalysisFile> static_files = new List<StaticAnalysisFile>();

        public static int lingqucishu = 2;
        public static int pianli_1 = 1;

        public static List<string> software_list = new List<string>();
        public static Dictionary<string, FileList> beicejianqingdan_dict = new Dictionary<string, FileList>();

        public static bool ceshidagang = true;
        public static int pianli_2 = 1;

        public static List<string> pingshenzuchengyuan = new List<string>();
        public static Dictionary<string, SoftwareItems> software_dict;
    }
}
