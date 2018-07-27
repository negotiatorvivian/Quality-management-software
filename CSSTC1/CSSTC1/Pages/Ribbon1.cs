using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Aspose.Words;

using CSSTC1.FileProcessors;
using CSSTC1.InputProcessors;

namespace CSSTC1.Pages {
    public partial class Ribbon1 {
        public FileReader1 file_reader = new FileReader1();
        public ProjectEstabInfoProcessor project_estab_info = new ProjectEstabInfoProcessor();

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog1.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog1.FileName;
                file_reader.read_charts(read_in_file);
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e) {
            this.project_estab_info.show_estab_info();
        }
    }
}
