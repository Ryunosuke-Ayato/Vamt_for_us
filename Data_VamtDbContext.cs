using Microsoft.EntityFrameworkCore;
using Vamt_for_us.Models;
using Vamt_for_us.Services;
namespace Vamt_for_us.Data
{
    public class VamtDbContext : DbContext
    {
        public VamtDbContext() { }
        public DbSet<ProductKey> ProductKeys { get; set; }
        public DbSet<ActiveProduct> ActiveProducts { get; set; }
        public DbSet<LicenseStatusText> LicenseStatusTexts { get; set; }
        public DbSet<ProductKeyTypeName> ProductKeyTypeNames { get; set; }
        public VamtDbContext(DbContextOptions<VamtDbContext> options) : base(options) { }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer(ConfigurationService.ConnectionString);
            }
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ProductKey>().ToTable("ProductKey", "base");
            modelBuilder.Entity<ActiveProduct>().ToTable("ActiveProduct", "base");
            modelBuilder.Entity<LicenseStatusText>().ToTable("LicenseStatusText", "base");
            modelBuilder.Entity<ProductKeyTypeName>().ToTable("ProductKeyTypeName", "base");
        }
    }
}
