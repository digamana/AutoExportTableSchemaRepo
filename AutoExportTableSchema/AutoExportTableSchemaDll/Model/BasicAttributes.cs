using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema.Model
{
    public class BasicAttributes
    {
        public string SchemaName { get; set; }
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public string ColumnType { get; set; }
        public string MaxLength { get; set; }
        public string IsNull { get; set; }
        public string IsPrimaryKey { get; set; }
        public string Describe { get; set; }
    }
}
