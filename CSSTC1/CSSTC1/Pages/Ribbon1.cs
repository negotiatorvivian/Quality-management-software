using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Aspose.Words;

using CSSTC1.FileProcessors;

namespace CSSTC1.Pages {
    public partial class Ribbon1 {
        public FileReader1 file_reader = new FileReader1();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e) {
            if(openFileDialog1.ShowDialog() == DialogResult.OK) {
                string read_in_file = openFileDialog1.FileName;
                file_reader.read_charts(read_in_file);
            }
        }
    }
}
