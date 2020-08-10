using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB.Entity
{
    public class Cell
    {
        public Cell()
        {

        }
        public Cell(int rowIndex,int colIndex, object value)
        {
            RowIndex = rowIndex;
            ColumnIndex = colIndex;
            Value = value;
        }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public object Value { get; set; }
    }
}
