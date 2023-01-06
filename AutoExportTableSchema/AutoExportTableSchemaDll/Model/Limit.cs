using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema.Model
{
    public class Limit
    {
        public int Min { get; set; }
        public int Max { get; set; }
        public Limit(int Min,int Max)
        {
            this.Min = Min;
            this.Max = Max;
        }
    }
}
