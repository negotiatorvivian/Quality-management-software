using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.ConstantVariables {
    class ContentFlags {
        public static bool peizhiceshi;
        public static bool xitongceshi;
        public static int missing = -1;

        public static Dictionary<string, int> file_type_counter = new Dictionary<string,int>();
    }
}
