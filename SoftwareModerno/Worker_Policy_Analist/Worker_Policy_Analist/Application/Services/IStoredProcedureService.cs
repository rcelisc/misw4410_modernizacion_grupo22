using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worker_Policy_Analist.Application.Services
{
    public interface IStoredProcedureService
    {
        Task ExecuteStoredProcedureAsync(bool blCompleto);
    }

}
