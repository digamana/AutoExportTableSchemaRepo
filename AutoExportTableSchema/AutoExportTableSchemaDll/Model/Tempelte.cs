using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExportTableSchema.Model
{
    public class Tempelte
    {
        public object ServerName { get; set; }
        public object DbName { get; set; }
        public object Account { get; set; }
        public object Password { get; set; }
        public Tempelte(object ServerName, object DbName, object Account, object Password)
        {
            this.ServerName = ServerName;
            this.DbName = DbName;
            this.Account = Account;
            this.Password = Password;

        }
    }
}
