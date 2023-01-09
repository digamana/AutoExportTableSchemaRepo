using AutoExportTableSchema.Model;
using AutoExportTableSchemaDll.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchemaDll.Export
{
    public interface IMappingExcelSkill
    {
        IEnumerable<BasicAttributes> GetExcelData();
        void SetExcelData(IEnumerable<BasicAttributes> lstTableStruct);
        void Save();
    }
}
