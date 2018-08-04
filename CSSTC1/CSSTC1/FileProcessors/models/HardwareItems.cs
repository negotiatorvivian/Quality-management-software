using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class HardwareItems {
        public string yj_mingcheng;
        public string yj_bianhao;
        public string yj_tongtu;
        public string yj_zhuangtai;
        public string wd_laiyuan;

        public HardwareItems(string yj_mingcheng, string yj_bianhao, string yj_tongtu, string yj_zhuangtai, 
            string wd_laiyuan) {
            this.yj_mingcheng = yj_mingcheng;
            this.yj_bianhao = yj_bianhao;
            this.yj_tongtu = yj_tongtu;
            this.yj_zhuangtai = yj_zhuangtai;
            this.wd_laiyuan = wd_laiyuan;
        }
    }
}
