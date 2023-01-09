using AutoExportTableSchema.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchemaDll.Model
{
    public class OutputExcelStruct
    {
        public string TableName { get; set; }
        public BasicAttributes BasicAttributes { get; set; }
    }
}
