using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class StaticAnalysisFile {
        public string rj_mingcheng;
        public string xt_banben;
        public string bc_yuyan;
        public string yz_danwei;
        public string jtfx_fanwei;

        public StaticAnalysisFile(string rj_mingcheng, string jtfx_fanwei, string xt_banben,
            string bc_yuyan, string yz_danwei) {
            this.rj_mingcheng = rj_mingcheng;
            this.xt_banben = xt_banben;
            this.bc_yuyan = bc_yuyan;
            this.yz_danwei = yz_danwei;
            this.jtfx_fanwei = jtfx_fanwei;
        }
    }
}
