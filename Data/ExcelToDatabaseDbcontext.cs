using ExcelToDatabase.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelToDatabase.Data
{
    public class ExcelToDatabaseDbcontext : DbContext
    {
        public ExcelToDatabaseDbcontext(DbContextOptions<ExcelToDatabaseDbcontext> options) : base(options) 
        {
            
        }
        public DbSet<StudentExcel> StudentExcels { get; set; }
    }
}
