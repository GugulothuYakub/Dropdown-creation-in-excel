using entitycore.Data;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace entitycore.Data
{
    public class SuperheroContext : DbContext
    {
        public SuperheroContext(DbContextOptions<SuperheroContext> options) : base(options)
        {

        }
        public DbSet<SuperHero> SuperHeroes { get; set; }
        public DbSet<Address> Address { get; set; }
        public DbSet<Movies> Movies { get; set; }
        public DbSet<ElmahError> ElmahErrors { get; set; }
    }
    public class SuperHero
    {
        [Key]
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Place { get; set; } = string.Empty;
        [ForeignKey("AddressId")]
        public int AddressId { get; set; }
        [ForeignKey("MoviesId")]
        public int MoviesId { get; set; }
        public virtual Address address { get; set; }
        public virtual Movies movies { get; set; }
    }
    public class Address
    {
        [Key]
        public int Id { get; set; }
        public string locality { get; set; } = string.Empty;
        public string City { get; set; } = string.Empty;
        public string State { get; set; } = string.Empty;
        public string PostalCode { get; set; } = string.Empty;
    }
    public class Movies
    {
        [Key]
        public int Id { get; set; }
        public int movieid { get; set; } = 0;
        public string moviename { get; set; } = string.Empty;
        public DateTime? date { get; set; }
    }

    public class ElmahError
    {
        [Key]
        public Guid ErrorId { get; set; }
        public string Application { get; set; } = string.Empty;
        public string Host { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string User { get; set; } = string.Empty;
        public int StatusCode { get; set; }
        public DateTime TimeUtc { get; set; }
        public int Sequence { get; set; }
        public string AllXml { get; set; } = string.Empty;
    }
}
