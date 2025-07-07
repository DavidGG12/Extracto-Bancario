using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Db.Interface;

namespace Template_Tesoreria.Helpers.Db
{
    public class DbFactory : IDatabaseManager
    {
        private readonly string _connectionString;

        public DbFactory(string connectionString)
        {
            _connectionString = connectionString;
        }

        public DataTable ExecuteStoredProcedure(string storedProcedureName, Dictionary<string, object> parameters)
        {
            DataTable resultTable = new DataTable();

            using (var con = new SqlConnection(_connectionString))
            {
                using(var cmd = new SqlCommand(storedProcedureName, con))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    if(parameters != null)
                    {
                        foreach(var param in parameters)
                        {
                            if (param.Value is string)
                                cmd.Parameters.AddWithValue(param.Key, param.Value.ToString().Replace("'", "") ?? "");
                            else
                                cmd.Parameters.AddWithValue(param.Key, param.Value);
                        }
                    }

                    con.Open();

                    using (var adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(resultTable);
                        return resultTable;
                    }
                }
            }
        }

        public class DatabaseManagerFactory
        {
            public static IDatabaseManager CreateDatabaseManager(string connectionString)
            {
                return new DbFactory(connectionString);
            }
        }
    }
}
