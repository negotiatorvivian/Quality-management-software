using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    class FileList {
        public string xm_biaoshi;
        public string wd_mingcheng;
        public string wd_banben;
        public string fb_riqi;
        public string wd_laiyuan;

        public FileList(string xm_biaoshi, string wd_mingcheng, string wd_banben,string fb_riqi,string wd_laiyuan) {
            this.xm_biaoshi = xm_biaoshi;
            this.wd_mingcheng = wd_mingcheng;
            this.wd_banben = wd_banben;
            this.fb_riqi = fb_riqi;
            this.wd_laiyuan = wd_laiyuan;
        }
    }
}
