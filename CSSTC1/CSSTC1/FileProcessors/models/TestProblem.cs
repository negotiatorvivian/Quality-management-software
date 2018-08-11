using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileProcessors.models {
    public class TestProblem {
        public string cs_stage;
        public List<int> zhimingwenti;
        public List<int> yanzhongwenti;
        public List<int> yibanwenti;
        public List<int> qingweiwenti;

        public TestProblem(string cs_stage, List<int> zhimingwenti, List<int> yanzhongwenti, List<int> yibanwenti,
            List<int> qingweiwenti) {
            this.cs_stage = cs_stage;
            this.zhimingwenti = zhimingwenti;
            this.yanzhongwenti = yanzhongwenti;
            this.yibanwenti = yibanwenti;
            this.qingweiwenti = qingweiwenti;
        }
    }
}
