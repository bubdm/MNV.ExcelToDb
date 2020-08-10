using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB
{
    public class MariaDbColumn
    {
        public UInt64 ColumnID { get; set; }
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public UInt64 MaxLength { get; set; }
        public Decimal DataPrecision { get; set; }
        public bool IsNullable { get; set; }
        public string ColumnDefault { get; set; }

        private string _excelColHeaderName;
        public string ExcelColHeaderName
        {
            get { return _excelColHeaderName; }
            set
            {
                //var excelColumn = ExcelColumn.excelColumns.FirstOrDefault(col => col.Name == value);
                //if (excelColumn != null)
                //{
                //    ExcelColPossition = excelColumn.Possiotion;
                //}
                _excelColHeaderName = value;
            }
        }

        private int? _excelColPossition;
        public int? ExcelColPossition {
            get { return _excelColPossition; }
            set
            {
                _excelColPossition = value;
                var excelColumn = ExcelColumn.excelColumns.FirstOrDefault(col => col.Possiotion == value);
                if (excelColumn != null)
                {
                    ExcelColHeaderName = excelColumn.Name;
                }
                else
                {
                    ExcelColHeaderName = string.Empty;
                }
            }
        }
    }
}
