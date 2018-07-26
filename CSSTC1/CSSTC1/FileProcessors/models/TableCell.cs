using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSSTC1.FileReaders {
    public class TableCell {
        public int tableIndex;
        public int rowIndex;
        public int cellIndex;
        public string celltext;

        public TableCell(int PiTableID, int PiRowID, int CellID, string text) {
            tableIndex = PiTableID;
            rowIndex = PiRowID;
            cellIndex = CellID;
            celltext = text;
        }
    }
}
