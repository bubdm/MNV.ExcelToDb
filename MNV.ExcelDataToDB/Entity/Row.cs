using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB.Entity
{
    public class Row
    {
        public Row()
        {

        }
        public Row(int rowIndex)
        {
            RowIndex = rowIndex;
        }
        public int RowIndex { get; set; }
        public List<Cell> Cells = new List<Cell>();
    }
}
