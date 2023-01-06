using AutoExportTableSchema.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema.SqlConnect
{
    public interface ISqlData
    {
        IEnumerable<BasicAttributes> lstAttributes();
    }
}
