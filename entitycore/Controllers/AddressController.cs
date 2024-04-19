using entitycore.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace entitycore.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AddressController : ControllerBase
    {
        private readonly SuperheroContext _context;
        private readonly IConfiguration _configuration;
        private readonly ILogger<AddressController> _logger;
        private readonly string _className;
        public AddressController(SuperheroContext context, IConfiguration configuration, ILogger<AddressController> logger)
        {
            _context = context;
            _configuration = configuration;
            _logger = logger;
            _className = context.GetType().Name;
        }
        [HttpGet]
        public async Task<ActionResult<List<Address>>> Get()
        {
            try
            {
                _logger.LogInformation("Entered into Getmethod: {ClassName}", _className);

                var AddrIds = await _context.Address.Select(a => a.Id).ToListAsync();

                // Create Excel file
                var filePath = @"C:\Users\yakub.gugulothu\Downloads\WSO\file.xlsx";
                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Address");

                    // Create a hidden worksheet
                    var hiddenWorksheet = package.Workbook.Worksheets.Add("Hidden");
                    hiddenWorksheet.Hidden = eWorkSheetHidden.Hidden;

                    // Write the address IDs to the hidden worksheet
                    for (int i = 0; i < AddrIds.Count; i++)
                    {
                        hiddenWorksheet.Cells[i + 1, 1].Value = AddrIds[i];
                    }

                    // Set data validation for each cell in the B column to use the range of cells with the address IDs from the hidden worksheet
                    for (int i = 1; i <= AddrIds.Count; i++)
                    {
                        var dropdownCell = worksheet.Cells[i, 2];
                        var dataValidation = dropdownCell.DataValidation.AddListDataValidation();
                        dataValidation.ShowErrorMessage = true;
                        dataValidation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                        dataValidation.ErrorTitle = "An invalid value was entered";
                        dataValidation.Error = "Please select a value from the list";

                        // Set the formula after setting all other properties
                        dataValidation.Formula.ExcelFormula = $"=Hidden!A1:A{AddrIds.Count}";
                    }

                    package.Save();
                }

                _logger.LogInformation("Exiting from Getmethod: {ClassName}", _className);

                return Ok(AddrIds);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error Occured in Getmethod: {ClassName}", _className);
                throw ex;
            }
        }

        // Method to split a list into chunks
        private List<List<T>> SplitListIntoChunks<T>(IEnumerable<T> list, int chunkSize)
        {
            var chunks = new List<List<T>>();
            var currentChunk = new List<T>();

            foreach (var item in list)
            {
                if (currentChunk.Count >= chunkSize)
                {
                    chunks.Add(currentChunk);
                    currentChunk = new List<T>();
                }

                currentChunk.Add(item);
            }

            if (currentChunk.Any())
            {
                chunks.Add(currentChunk);
            }

            return chunks;
        }

        [HttpPost]
        public async Task<ActionResult<Address>> AddAddress(Address AddAdress)
        {
            if (AddAdress == null)
                return BadRequest();

            _context.Address.Add(AddAdress);
            await _context.SaveChangesAsync();
            var addr = _context.Address.FindAsync(AddAdress.Id);

            return Ok(addr);
        }
    }
}
