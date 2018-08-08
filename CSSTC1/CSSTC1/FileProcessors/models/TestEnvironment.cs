using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class TestEnvironment {
        public string rj_mingcheng;
        public string diyici_ceshi;
        public string dierci_ceshi;
        public string ceshi_didian1;
        public string ceshi_didian2;

    public TestEnvironment(string rj_mingcheng, string diyici_ceshi, string dierci_ceshi, string ceshi_didian1,
        string ceshi_didian2) {
        this.rj_mingcheng = rj_mingcheng;
        this.diyici_ceshi = diyici_ceshi;
        this.dierci_ceshi = dierci_ceshi;
        this.ceshi_didian1 = ceshi_didian1;
        this.ceshi_didian2 = ceshi_didian2;
        }
    }

}
