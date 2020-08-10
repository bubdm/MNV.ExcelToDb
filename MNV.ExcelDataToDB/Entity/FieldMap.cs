using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB
{
    public class FieldMap
    {
        public FieldMap(string dbFieldName, string excelFieldName, int excelColPossition)
        {
            this.DbFieldName = dbFieldName;
            this.ExcelFieldName = excelFieldName;
            this.ExcelColPossition = excelColPossition;
        }
        public string DbFieldName { get; set; }
        public string ExcelFieldName { get; set; }
        public int ExcelColPossition { get; set; }
    }
}
