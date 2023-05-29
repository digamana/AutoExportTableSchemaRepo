using AutoExportTableSchema;
using AutoExportTableSchema.Domain;
using AutoExportTableSchema.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchemaDll.Domain
{
    public class Reading
    {
        public Excel _excel;
        public string _strSource { get; set; }
        public string _strTarget { get; set; }
        List<BasicAttributes> basicAttributes { get; set; }
        public Reading(string strSource)
        {
            _strSource = strSource;
            basicAttributes = new List<BasicAttributes>();
        }
        public void Run() 
        {
            basicAttributes.Clear();
            _excel = new Excel(_strSource, "總表");
            for (int i = 0; i < _excel.sheet.Dimension.End.Row; i++)
            {
 
                var E1 = _excel.sheet.Cells[i + 1, 1].Value;
                var E2 = _excel.sheet.Cells[i + 1, 2].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 2].Value.ToString(); ;
                var E3 = _excel.sheet.Cells[i + 1, 3].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 3].Value.ToString(); ;
                string E8 = _excel.sheet.Cells[i + 1, 8].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 8].Value.ToString();
                string E9 = _excel.sheet.Cells[i + 1, 9].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 9].Value.ToString();
                if (!string.IsNullOrEmpty(E2) && !string.IsNullOrEmpty(E3) && !E1.Equals("SchemaName"))
                {
                    basicAttributes.Add(new BasicAttributes()
                    {
                        SchemaName= E1.ToString(),
                        TableName = E2.ToString(),
                        ColumnName = E3.ToString(),
                        Describe = E8,
                        Describe2 = E9,
                    });
                }
            }
            basicAttributes = basicAttributes.Where(
                c =>   c.Describe   != ""  &&
                       c.SchemaName != ""  &&
                       c.TableName  != ""  &&
                       c.ColumnName != "" 
                ).ToList();
            string strPath = Path.Combine(Center.Downloads, $"CommandSQL_{DateTime.Now.ToString("yymmddhhmmss")}.sql");
            using (StreamWriter writer = new StreamWriter(strPath))
            {
                foreach (var item in basicAttributes)
                {
                    writer.WriteLine($"EXEC sp_addextendedproperty 'MS_Description', '{item.Describe}', 'SCHEMA', '{item.SchemaName}', 'TABLE', '{item.TableName}', 'COLUMN', '{item.ColumnName}';");
                }
            }
            Process.Start(strPath);
            
        }
    }
}
