using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MNV.ExcelDataToDB
{
    public class Field
    {
        public Field()
        {
            
        }
        public Field(string name, object value)
        {
            FieldName = name;
            FieldValue = value;
        }
        public string FieldName { get; set; }
        public object FieldValue { get; set; }
    }
}
