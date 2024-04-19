global using Microsoft.EntityFrameworkCore;
using entitycore;
using entitycore.Data;
using entitycore.MiddleWare;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using Serilog;
using Serilog.Events;

var builder = WebApplication.CreateBuilder(args);

// Set the license context
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddLogging(logging =>
{
    var config = builder.Configuration.GetSection("Logging");
    logging.AddFilter("Microsoft.EntityFrameworkCore.Database.Command", LogLevel.Warning);
    logging.AddFilter("Microsoft.EntityFrameworkCore.Infrastructure", LogLevel.Warning);
    logging.AddFilter("Microsoft.EntityFrameworkCore.Model.Validation", LogLevel.None);
    logging.AddFilter("Microsoft.EntityFrameworkCore.Model.Validation.ValidationLogger", LogLevel.None);

    logging.AddProvider(new FileLoggerProvider(builder.Configuration.GetSection("LogDirectory").Value.ToString()));
    logging.SetMinimumLevel(LogLevel.Information);
});
builder.Services.AddDbContext<SuperheroContext>(options =>
{
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
});
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
    app.UseDeveloperExceptionPage();
}

app.UseAuthorization();

app.UseMiddleware(typeof(ExceptionHandlingMiddleware));

app.MapControllers();

app.Run();