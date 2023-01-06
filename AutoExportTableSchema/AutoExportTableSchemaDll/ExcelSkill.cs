using AutoExportTableSchema.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema 
{
    public class ExcelSkill : IExcelSkill
    {
        ExcelWorksheet workSheet { get; set; }
        public ExcelSkill(ExcelWorksheet workSheet)
        {
            this.workSheet = workSheet;
        }
        public IEnumerable<int> getMergeRoews()
        {
            List<int> tempLst = new List<int>();
            for (int i = 1; i < workSheet.Dimension.End.Row; i++)
            {
                if (workSheet.Cells[$"A{i}:H{i}"].Merge == true)
                {
                    tempLst.Add(i);                 
                }
             
            }
            return tempLst;
        }
        public void setMergeValue(List<string> list)
        { 
            var temp = getMergeRoews();
            int i =0;
            foreach (var item in temp)
            {
                workSheet.Cells[$"A{item}"].Value = list[i];
                if (i < list.Count() - 1) i++;
                else break;
            }
        }
        public IEnumerable<Limit> getMergeRange(out int last)
        {
            var lst = getMergeRoews().ToList();
            var resLst = new List<Limit>();
            for (int ii = 0; ii < lst.Count() - 1; ii++)
            {
                var a1 = lst[ii] + 1;
                var a2 = lst[ii + 1] - 1;

                resLst.Add(new Limit(lst[ii] + 1, lst[ii + 1] - 1));
            }
            last = lst[getMergeRoews().Count() - 1];
            return resLst;
        }
    }
}
