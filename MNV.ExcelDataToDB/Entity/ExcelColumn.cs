using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB
{
    public class ExcelColumn
    {
        public static List<ExcelColumn> excelColumns = new List<ExcelColumn>();
        public ExcelColumn()
        {

        }
        public ExcelColumn(string name, int possition)
        {
            Name = name;
            Possiotion = possition;
        }
        public string Name { get; set; }
        public int Possiotion { get; set; }
    }
}
