using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class QestionReport {
        public string wenti;
        public string chulicuoshi;
        public string tiwendanwei;
        public string beizhu;

        public QestionReport(string wenti, string chulicuoshi, string tiwendanwei, string beizhu) {
            this.wenti = wenti;
            this.chulicuoshi = chulicuoshi;
            this.tiwendanwei = tiwendanwei;
            this.beizhu = beizhu;
        }

    }
}
