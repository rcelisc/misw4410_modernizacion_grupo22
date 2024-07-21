using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worker_Policy_Analist.Infrastructure.Data
{
    public interface IDatabaseConnectionService
    {
        DbConnection GetConnection();
    }
}
