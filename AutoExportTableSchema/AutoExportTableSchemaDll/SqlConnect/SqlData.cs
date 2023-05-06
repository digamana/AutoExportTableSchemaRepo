using AutoExportTableSchema.Model;
using Dapper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema.SqlConnect
{
    public class SqlData : ISqlData
    {
        /// <summary>
        /// 備註 要變更DB 請變更ConnectString的字串
        /// </summary>
         
        public string CONNECTION_STRING { get; set; }
        public SqlData(string CONNECTION_STRING)
        {
            this.CONNECTION_STRING = CONNECTION_STRING;
        }
        public IEnumerable<BasicAttributes> lstAttributes()
        {
            using (SqlConnection conn = new SqlConnection(this.CONNECTION_STRING))
            {
                string strSql = @"
 with main as
(
SELECT
c.name 'Column Name',
t.Name 'Column Type',
c.max_length 'Max Length',
case
when c.is_nullable = 0 then 'Not Null'
when c.is_nullable = 1 then 'Is Null'
End as 'IsNull',
ISNULL(i.is_primary_key, 0) 'IsPrimaryKey',
c.object_id 'object_id'
FROM
sys.columns c
INNER JOIN
sys.types t ON c.user_type_id = t.user_type_id
LEFT OUTER JOIN
sys.index_columns ic ON ic.object_id = c.object_id AND ic.column_id = c.column_id
LEFT OUTER JOIN
sys.indexes i ON ic.object_id = i.object_id AND ic.index_id = i.index_id
)
SELECT
s.name AS SchemaName,
t.name AS TableName,
main.[Column Name] as N'ColumnName',
main.[Column Type] as N'ColumnType',
main.[Max Length] as N'MaxLength',
main.[IsNull],
main.[IsPrimaryKey],
sys.extended_properties.value as N'ColumnDescription'
FROM sys.tables t
INNER JOIN sys.schemas s
ON t.schema_id = s.schema_id
left join main on main.object_id = OBJECT_ID(t.name　)
left join sys.extended_properties on sys.extended_properties.major_id = main.object_id and sys.extended_properties.minor_id = columnproperty(main.object_id, main.[Column Name], 'ColumnId') and sys.extended_properties.name = 'MS_Description'
";
                var accounts = conn.Query<BasicAttributes>(strSql);
                return accounts;
            }
        }
    }
}
