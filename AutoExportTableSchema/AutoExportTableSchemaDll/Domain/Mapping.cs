using AutoExportTableSchemaDll.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchemaDll.Domain
{
    public class Mapping
    {
        public string _strSource { get; set; }
        public string _strTarget { get; set; }
        public Mapping(string strSource,string strTarget) 
        {
            _strSource=strSource;
            _strTarget = strTarget;
        }
        public void Run()
        {
            IMappingExcelSkill SourceExcel = new MappingExcelSkill(@"C:\Users\User\Downloads\新增資料夾\DemoSource.xlsx");
            var temp=SourceExcel.GetExcelData();
            IMappingExcelSkill SourceExce2 = new MappingExcelSkill(@"C:\Users\User\Downloads\新增資料夾\DemoTarget.xlsx");
            SourceExce2.SetExcelData(temp);
            SourceExce2.Save();
            //IMappingExcelSkill TargetExcel = new MappingExcelSkill();
        }
    }
}
