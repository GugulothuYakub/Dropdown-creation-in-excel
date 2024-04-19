using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using entitycore.Data;
using entitycore.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Serilog;
using System.Text;

namespace entitycore.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SuperHeroController : ControllerBase
    {
        private readonly SuperheroContext _context;
        private readonly ILogger<SuperHeroController> _logger;
        public SuperHeroController(SuperheroContext context, ILogger<SuperHeroController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet("Test")]
        public IActionResult DownloadLogFile()
        {
            try
            {
                // Code that may throw an exception
                throw new Exception("Something went wrong");
            }
            catch (Exception ex)
            {
                //Log.Error(ex, "An error occurred");
                ////Log.Information(ex, "An error occurred");
                //Log.Information("Additional information: This is an informational log message.");

                //// Get the class name dynamically
                //var className = this.GetType().Name;

                //Log.ForContext("ClassName", className)
                //    .ForContext("ActionName", this.ControllerContext.ActionDescriptor.ActionName)
                //    .Error(ex, "An error occurred");

                //// Get the controller and action names
                ////var controllerName = this.ControllerContext.ActionDescriptor.ControllerName;
                //var actionName = this.ControllerContext.ActionDescriptor.ActionName;

                //// Generate the log file name
                ////var logFileName = $"{controllerName}_{actionName}_error.log";
                //// Generate the log file name
                //var logFileName = $"{className}_error.log";
                //var logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "logs", logFileName);
                ////var logFilePath = Path.Combine(Directory.GetCurrentDirectory(), "logs", logFileName);

                //// Get the log file contents
                //var logFileContents = System.IO.File.ReadAllText(logFilePath);
                //var contentType = "text/plain";

                //return File(Encoding.UTF8.GetBytes(logFileContents), contentType, logFileName);

                //_logger.LogError(ex, "An error occurred");
                //_logger.LogInformation(ex, "An error occurred");

                //var className = this.GetType().Name;
                //_logger.ForContext("ClassName", className)
                //    //.ForContext("ActionName", this.ControllerContext.ActionDescriptor.ActionName)
                //    .Error(ex, "An error occurred");

                _logger.LogError(ex, "An error occurred");
                _logger.LogInformation(ex, "An error occurred");

                var className = this.GetType().Name;
                Log.ForContext("ClassName", className)
                    .Error(ex, "An error occurred");

                return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while processing your request.");
            }
        }

        [HttpGet("Test/{id}")]
        public IActionResult GetAll(int id)
        {
            try
            {
                _logger.LogInformation("Entered Superhero Get method");
                _logger.LogInformation($"Id value is: {id}");

                // Your Entity Framework Core logic here to fetch superheroes

                _logger.LogInformation("Successfully fetched superheroes");

                return Ok(/* superheroes */);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching superheroes");
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public async Task<ActionResult<List<SuperHero>>> Get()
        {
            try
            {
                _logger.LogInformation("entered superhero get method");


                //int a = 10;
                //_logger.LogInformation("a value is:" +a);

                //int b = 0;
                //_logger.LogInformation("b value is:" + b);

                //int c = a / b;
                //_logger.LogInformation("c value is:" + c);

                var _heroes = await _context.SuperHeroes.Include(x => x.address).Include(x => x.movies).ToListAsync();

                _logger.LogInformation("Successfully fetched superheroes");

                return Ok(_heroes);
            }
            catch (Exception ex)
            {
                //return BadRequest(ex.Message);
                //Log.Error(ex, "An error occurred"); // Log the exception with Serilog
                //return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while processing your request.");

                //Log.ForContext("ControllerName", this.ControllerContext.ActionDescriptor.ControllerName)
                // Get the class name dynamically
                var className = this.GetType().Name;

                Log.ForContext("ClassName", className)
                    //.ForContext("ActionName", this.ControllerContext.ActionDescriptor.ActionName)
                    .Error(ex, "An error occurred");
                return StatusCode(StatusCodes.Status500InternalServerError, "An error occurred while processing your request.");
            }
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<SuperHero>> Get(int id)
        {
            var hero = await _context.SuperHeroes.FindAsync(id);
            if (hero == null)
                return BadRequest("hero is not found");
            return Ok(hero);
        }
        //[HttpGet("{id}/Testing")]
        //public async Task<ActionResult<SuperHero>> Gets(int id)
        //{
        //    var hero = await _context.SuperHeroes
        //    .Where(s => s.Id == id).Select(s => new SuperHero { Id = s.Id }).FirstOrDefaultAsync();
        //    if (hero == null)
        //        return BadRequest("hero is not found");
        //    return Ok(hero);
        //}
        [HttpGet("Testing")]
        public async Task<IActionResult> Gets()
        {
            var heroIds = await _context.SuperHeroes
                .Select(s => s.Id)
                .ToListAsync();
            //if (!heroIds.Any())
            //    return BadRequest("No SuperHeroes found");

            

            return Ok(heroIds);
        }

        [HttpPost]
        public async Task<ActionResult<SuperHero>> AddHero(SuperHeroInputModel hero)
        {
            if (hero == null)
                return BadRequest();

            var _hero = new SuperHero
            {
                Name = hero.Name,
                FirstName = hero.FirstName,
                LastName = hero.LastName,
                Place = hero.Place,
                AddressId = hero.AddressId,
                MoviesId = hero.MoviesId
            };

            _context.SuperHeroes.Add(_hero);
            await _context.SaveChangesAsync();

            // Retrieve the inserted superhero by its Id
            //var insertedHero = await _context.SuperHeroes.FindAsync(_hero.Id);
            // Eagerly load the related Address and Movies
            await _context.Entry(_hero)
                .Reference(sh => sh.address)
                .LoadAsync();

            await _context.Entry(_hero)
                .Reference(sh => sh.movies)
                .LoadAsync();
            // Return the inserted superhero
            return Ok(_hero);
        }
        [HttpPut]
        public async Task<ActionResult<SuperHero>> UpdateHero(SuperHero request)
        {
            var dbHero = await _context.SuperHeroes.FindAsync(request.Id);
            if (dbHero == null)
                return BadRequest("hero is not found");

            dbHero.Name = request.Name;
            dbHero.FirstName = request.FirstName;
            dbHero.LastName = request.LastName;
            dbHero.Place = request.Place;

            await _context.SaveChangesAsync();
            // Eagerly load the related Address and Movies
            //await _context.Entry(dbHero)
            //    .Reference(sh => sh.address)
            //    .LoadAsync();

            //await _context.Entry(dbHero)
            //    .Reference(sh => sh.movies)
            //    .LoadAsync();
            return Ok(await _context.SuperHeroes.ToListAsync());
        }
        [HttpDelete("{Id}")]
        public async Task<ActionResult<SuperHero>> Delete(int Id)
        {
            var dbHero = await _context.SuperHeroes.FindAsync(Id);

            if (dbHero == null)
                return BadRequest("hero is not found");

            _context.SuperHeroes.Remove(dbHero);
            await _context.SaveChangesAsync();

            // Eagerly load the related Address and Movies
            await _context.Entry(dbHero)
                .Reference(sh => sh.address)
                .LoadAsync();

            await _context.Entry(dbHero)
                .Reference(sh => sh.movies)
                .LoadAsync();

            return Ok(await _context.SuperHeroes.ToListAsync());
        }
    }
}
