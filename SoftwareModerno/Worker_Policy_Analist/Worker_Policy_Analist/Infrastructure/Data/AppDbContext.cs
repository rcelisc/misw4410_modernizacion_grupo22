using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Worker_Policy_Analist.Domain.Entities;

namespace Worker_Policy_Analist.Infrastructure.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

        public DbSet<ItemPoliticasBasicas> ItemsPoliticasBasicas { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<ItemPoliticasBasicas>(entity =>
            {
                entity.HasNoKey(); // Configura la entidad como sin clave primaria
            });
        }

    }
}
