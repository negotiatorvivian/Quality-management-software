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

        public static HashSet<string> test_types= new HashSet<string>();
        public static int lingqucishu = 2;
        public static List<string> lingqushijian = new List<string>();
        public static string xiangmukaishishijian = "";
        public static int pianli_1 = 1;

        public static List<FileList> beicejianqingdan = new List<FileList>();
        public static List<string> content_list = new List<string>();
        public static bool ceshidagang = true;
    }
}
