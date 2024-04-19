using System.ComponentModel.DataAnnotations;

namespace entitycore.Model
{
    public class SuperHeroInputModel
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Place { get; set; } = string.Empty;
        public int AddressId { get; set; }
        public int MoviesId { get; set; }
    }
}
