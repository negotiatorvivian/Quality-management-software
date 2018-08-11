using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class TestExample {
        public string yl_type;
        public string sjyl_num;
        public string sjyl_percent;
        public string zxyl_num;
        public string wtg_num;

        public TestExample(string yl_type, string sjyl_num, string sjyl_percent, string zxyl_num,
            string wtg_num) {
            this.yl_type = yl_type;
            this.sjyl_num = sjyl_num;
            this.sjyl_percent = sjyl_percent;
            this.zxyl_num = zxyl_num;
            this.wtg_num = wtg_num;
        }
    }
}
