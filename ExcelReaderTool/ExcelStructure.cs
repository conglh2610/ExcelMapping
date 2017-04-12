using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTool
{
    public class ExcelStructure
    {
        public ExcelStructure(string col, string typ)
        {
            ColumnName = col;
            ColumnType = typ;
        }
        public String ColumnName { get; set; }
        public String ColumnType { get; set; }
    }
}
