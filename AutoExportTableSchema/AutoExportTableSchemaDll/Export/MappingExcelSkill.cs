using AutoExportTableSchema;
using AutoExportTableSchema.Model;
using AutoExportTableSchemaDll.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchemaDll.Export
{
    public class MappingExcelSkill: IMappingExcelSkill
    {
        public Excel _excel;
        public IEnumerable<BasicAttributes> _lstTableStruct { get; set; }
        public MappingExcelSkill(string strFilePath) 
        {
            _excel = new Excel(strFilePath, "總表");
            _lstTableStruct = GetExcelData();
        }
        public void SetExcelData(IEnumerable<BasicAttributes> lstTableStruct)
        {
            for (int i = 0; i < _excel.sheet.Dimension.End.Row; i++)
            {
                var E1 = _excel.sheet.Cells[i + 1, 1].Value;
                var E2 = _excel.sheet.Cells[i + 1, 2].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 2].Value.ToString(); ;
                var E3 = _excel.sheet.Cells[i + 1, 3].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 3].Value.ToString(); ;
                string E8 = _excel.sheet.Cells[i + 1, 8].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 8].Value.ToString();
                if (!string.IsNullOrEmpty(E2) && !string.IsNullOrEmpty(E3) && !E1.Equals("SchemaName"))
                {
                    string tempValue = GetDescribe(lstTableStruct, E2, E3);
                    _excel.sheet.Cells[i + 1, 8].Value = tempValue;
                }
            }
        }
        public void Save()
        {
            _excel.workbook.Save();
            _excel.Dispose();
        }
        public IEnumerable<BasicAttributes> GetExcelData()
        {
            List<BasicAttributes> output = new List<BasicAttributes>();
            for (int i = 0; i < _excel.sheet.Dimension.End.Row; i++)
            {
                var E1 = _excel.sheet.Cells[i + 1 ,1].Value;
                var E2 = _excel.sheet.Cells[i + 1, 2].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 2].Value.ToString(); ;
                var E3=  _excel.sheet.Cells[i + 1, 3].Value == null ? string.Empty : _excel.sheet.Cells[i + 1, 3].Value.ToString(); ;
                string E8 = _excel.sheet.Cells[i + 1, 8].Value ==null? string.Empty: _excel.sheet.Cells[i + 1, 8].Value.ToString();
                if ( !string.IsNullOrEmpty(E2) && !string.IsNullOrEmpty(E3)    && !E1.Equals("SchemaName") )
                {
                    output.Add(new BasicAttributes()
                    {
                    TableName=E2.ToString(),
                    ColumnName=E3.ToString(),
                    Describe = E8,
                    });
                }
            }
            return output;
        }
        public static string  GetDescribe(IEnumerable<BasicAttributes> lstTableStruct,string strTableName,string ColName)
        {
            var temp = lstTableStruct.Where(c => c.TableName == strTableName && c.ColumnName == ColName).FirstOrDefault();
            if (temp == null) return string.Empty;
            var temp2 = string.IsNullOrEmpty(temp.Describe) ? string.Empty: temp.Describe;
            return temp2;
        }

    }
}
