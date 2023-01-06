using AutoExportTableSchema.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema
{
    public interface IExcelSkill
    {
          IEnumerable<int> getMergeRoews();
          IEnumerable<Limit> getMergeRange(out int last);
          void setMergeValue(List<string> list);

    
    }
}
