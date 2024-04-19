using entitycore.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace entitycore.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MovieListController : ControllerBase
    {
        private readonly SuperheroContext _context;
        public MovieListController(SuperheroContext context)
        {
            _context = context;
        }
        [HttpGet]
        public async Task<ActionResult<List<Movies>>> Get()
        {
            var _movie = await _context.Movies.ToListAsync();

            return Ok(_movie);
        }
        [HttpPost]
        public async Task<ActionResult<Movies>> AddMovie(Movies AddMovies)
        {
            if (AddMovies == null)
                return BadRequest();

            var movie = _context.Movies.Add(AddMovies);
            await _context.SaveChangesAsync();

            // Retrieve the inserted movies by its Id
            var insertedMovie = await _context.Movies.FindAsync(AddMovies.Id);

            return Ok(insertedMovie);
        }
    }
}
