using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class SoftwareItems {
        public string rj_mingcheng;
        public string rj_banben;
        public string rj_yongtu;
        public string wd_laiyuan;
        public string rj_bianhao;

        public SoftwareItems(string rj_mingcheng, string rj_banben, string rj_yongtu, string wd_laiyuan,
            string rj_bianhao) {
            this.rj_mingcheng = rj_mingcheng;
            this.rj_banben = rj_banben;
            this.rj_yongtu = rj_yongtu;
            this.wd_laiyuan = wd_laiyuan;
            this.rj_bianhao = rj_bianhao;
        }
    }
}
