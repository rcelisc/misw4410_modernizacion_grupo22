using Microsoft.EntityFrameworkCore;
using System.Data;
using Microsoft.Data.SqlClient;
using Worker_Policy_Analist.Application.Services;
using static System.Formats.Asn1.AsnWriter;
using Microsoft.Extensions.Configuration;

namespace Worker_Policy_Analist
{
    public class Worker : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<Worker> _logger;
        private readonly IConfiguration _configuration;
        private bool _blCompleto;

        public Worker(IServiceProvider serviceProvider, ILogger<Worker> logger, IConfiguration configuration)
        {
            _serviceProvider = serviceProvider;
            _logger = logger;
            _configuration = configuration;

            // Leer la configuración desde appsettings.json
            _blCompleto = _configuration.GetValue<bool>("AppSettings:BlCompleto");
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                if (_logger.IsEnabled(LogLevel.Information))
                {
                    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                }
                try
                {
                    using (var scope = _serviceProvider.CreateScope())
                    {
                        var storedProcedureService = scope.ServiceProvider.GetRequiredService<IStoredProcedureService>();
                        await storedProcedureService.ExecuteStoredProcedureAsync(_blCompleto);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "An error occurred while executing the stored procedure.");
                }

                await Task.Delay(1000, stoppingToken);
            }
        }

    }
}