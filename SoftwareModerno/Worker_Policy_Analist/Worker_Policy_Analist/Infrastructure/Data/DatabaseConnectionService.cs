using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worker_Policy_Analist.Infrastructure.Data;

namespace Worker_Policy_Analist.Application.Services
{

    public class DatabaseConnectionService : IDatabaseConnectionService
    {
        private readonly string _connectionString;

        public DatabaseConnectionService(string connectionString)
        {
            _connectionString = connectionString;
        }

        public DbConnection GetConnection()
        {
            return new SqlConnection(_connectionString);
        }
    }
}
