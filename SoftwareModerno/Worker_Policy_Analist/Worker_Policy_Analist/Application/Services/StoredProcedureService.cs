using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worker_Policy_Analist.Domain.Entities;
using Worker_Policy_Analist.Infrastructure.Data;

namespace Worker_Policy_Analist.Application.Services
{
    public class StoredProcedureService : IStoredProcedureService
    {
        private readonly AppDbContext _context;
        private readonly ILogger<StoredProcedureService> _logger;

        public StoredProcedureService(AppDbContext context, ILogger<StoredProcedureService> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task ExecuteStoredProcedureAsync(bool blCompleto)
        {
            try
            {
                var items = await GetItemsPoliticasBasicasAsync(blCompleto);

                _logger.LogInformation($"Total Records Obtained: {items.Count}");

                foreach (var item in items)
                {
                    await ExecutePolicyProcedure(item);
                    _logger.LogInformation($"Processed Record: {item.NoSolicitud}");
                }

                _logger.LogInformation($"Total Records Processed: {items.Count}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing stored procedures.");
                throw; // Re-throw the exception to let it propagate if necessary
            }
        }

        private async Task<List<ItemPoliticasBasicas>> GetItemsPoliticasBasicasAsync(bool blCompleto)
        {
            try
            {

                var items = new List<ItemPoliticasBasicas>();
                string sql = "EXEC SP_GET_ITEMS_POLITICAS_BASICAS @blCompleto";
                var param = new SqlParameter("@blCompleto", blCompleto);

                items = await _context.ItemsPoliticasBasicas.FromSqlRaw(sql, param).ToListAsync();
                return items;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing stored procedure SP_GET_ITEMS_POLITICAS_BASICAS.");
                throw;
            }
        }


        private async Task ExecutePolicyProcedure(ItemPoliticasBasicas item)
        {
            try
            {
                var parameters = new[]
                {
                    new SqlParameter("@NoSolicitud", item.NoSolicitud),
                    new SqlParameter("@Tipo", item.TipoId),
                    new SqlParameter("@Linea", item.NumId),
                };

                string sql = "EXEC SP_POLITICAS_BASICAS @NoSolicitud, @Tipo, @NumId, @Linea";
                await _context.Database.ExecuteSqlRawAsync(sql, parameters);

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error executing stored procedure SP_POLITICAS_BASICAS.");
                throw;
            }
        }
    }
}

//public async Task ExecuteStoredProcedureAsync1(bool blCompleto)
//        {
//            try
//            {
//                using var connection = _databaseConnectionService.GetConnection();
//                using var command = new SqlCommand("SP_POLITICAS_BASICAS", (SqlConnection)connection)
//                {
//                    CommandType = CommandType.StoredProcedure
//                };

//                // Agregar parámetros.
//                command.Parameters.AddWithValue("@NoSolicitud", noSolicitud);
//                command.Parameters.AddWithValue("@Tipo", tipo);
//                command.Parameters.AddWithValue("@Linea", linea);

//                await connection.OpenAsync();
//                await command.ExecuteNonQueryAsync();
//            }
//            catch (Exception ex)
//            {
//                _logger.LogError(ex, "Error executing stored procedure.");
//                throw; // Re-throw the exception to let it propagate if necessary
//            }
//        }
//    }
//}
