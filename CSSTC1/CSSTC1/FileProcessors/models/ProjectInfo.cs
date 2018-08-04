using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    class ProjectInfo {
        public string xt_mingcheng;
        public string rj_mingcheng;
        public string xt_banben;
        public string yx_huanjing;
        public string kf_huanjing;
        public string bc_yuyan;
        public string dm_guimo;
        public string dengji;
        public string yz_danwei;
        public string beizhu;

        public ProjectInfo(string xt_mingcheng, string rj_mingcheng, string xt_banben, string yx_huanjing, 
            string kf_huanjing, string bc_yuyan, string dm_guimo, string dengji, string yz_danwei, 
            string beizhu) {
            this.xt_mingcheng = xt_mingcheng;
            this.rj_mingcheng = rj_mingcheng;
            this.xt_banben = xt_banben;
            this.yx_huanjing = yx_huanjing;
            this.kf_huanjing = kf_huanjing;
            this.bc_yuyan = bc_yuyan;
            this.dm_guimo = dm_guimo;
            this.dengji = dengji;
            this.yz_danwei = yz_danwei;
            this.beizhu = beizhu;
        }

        public ProjectInfo() {
            // TODO: Complete member initialization
            this.beizhu = "";
        }
    }
}
