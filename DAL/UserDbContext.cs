using AdminPanelCrud.Models;
using Microsoft.EntityFrameworkCore;


namespace AdminPanelCrud.DAL
{
    public class UserDbContext : DbContext
    {
        public UserDbContext (DbContextOptions<UserDbContext> options) : base(options) {
            
        }

        public DbSet<User> User { get; set; }
        public DbSet<Product> Product { get; set; }
    }
}
