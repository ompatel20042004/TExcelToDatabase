using ExcelToDatabase.Data;
using ExcelToDatabase.Uploads;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

//builder.Services.AddControllers();
builder.Services.AddControllers()
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.ReferenceHandler = System.Text.Json.Serialization.ReferenceHandler.Preserve;
    });

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.OperationFilter<SwaggerFileUploadOperationFilter>(); 
});


//builder.Services.AddDbContext<ExcelToDatabaseDbcontext>(dbContextOptions =>
//{
//    dbContextOptions.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
//    .EnableSensitiveDataLogging()
//    .EnableDetailedErrors();

//});
//var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<ExcelToDatabaseDbcontext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
           .LogTo(Console.WriteLine, LogLevel.Information));

var app = builder.Build();



if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage(); // Shows the actual error details
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseRouting();
app.UseAuthorization();
app.MapControllers();
app.UseHttpsRedirection();
app.Run();
