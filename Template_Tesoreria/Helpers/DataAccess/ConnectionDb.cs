using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.DataAccess
{
    public class ConnectionDb
    {
        public string DbTesoreria1019()
        {
            var server = System.Configuration.ConfigurationManager.AppSettings["CnnServer"] as string;
            var bd = System.Configuration.ConfigurationManager.AppSettings["CnnBD"] as string;
            var user = System.Configuration.ConfigurationManager.AppSettings["CnnUser"] as string;
            var pass = System.Configuration.ConfigurationManager.AppSettings["CnnPwd"] as string;
            var cnn = string.Format("Data Source={0};Initial Catalog={1};Perist Security Info=True;User ID={2};Password={3}", server, bd, user, pass);
            return cnn;
        }
    }
}
