using Worker_Policy_Analist;
using Microsoft.EntityFrameworkCore;
using Worker_Policy_Analist.Infrastructure.Data;
using Worker_Policy_Analist.Application.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var builder = Host.CreateApplicationBuilder(args);

// Configuración de la lectura del archivo de configuración
builder.Configuration.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

// Configuración del DbContext
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DataBaseConnection")));

// Registro del servicio de procedimiento almacenado
builder.Services.AddTransient<IStoredProcedureService, StoredProcedureService>();

// Registro del servicio alojado (worker)
builder.Services.AddHostedService<Worker>();

try
{
    var host = builder.Build();
    host.Run();
}
catch (Exception ex)
{
    Console.WriteLine($"Error al iniciar la aplicación: {ex.Message}");
    throw;
}
