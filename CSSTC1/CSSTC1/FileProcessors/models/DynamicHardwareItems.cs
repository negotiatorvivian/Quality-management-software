using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class DynamicHardwareItems {
        public string yj_mingcheng;
        public string yj_bianhao;
        public string yj_shuliang;
        public string yj_yongtu;
        public string yj_peizhi;
        public string yj_zhuangtai;
        public string wd_laiyuan;

        public DynamicHardwareItems(string yj_mingcheng, string yj_bianhao, string yj_shuliang, 
            string yj_yongtu, string yj_peizhi, string yj_zhuangtai, string wd_laiyuan) {
            this.yj_mingcheng = yj_mingcheng;
            this.yj_bianhao = yj_bianhao;
            this.yj_shuliang = yj_shuliang;
            this.yj_yongtu = yj_yongtu;
            this.yj_peizhi = yj_peizhi;
            this.yj_zhuangtai = yj_zhuangtai;
            this.wd_laiyuan = wd_laiyuan;
        }
    }
}
