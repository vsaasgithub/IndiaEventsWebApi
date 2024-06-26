
using DinkToPdf.Contracts;
using DinkToPdf;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.IdentityModel.Tokens;
using Serilog;
using System.Text;
using Microsoft.AspNetCore.Http.Headers;
using Microsoft.Extensions.DependencyInjection;

using Microsoft.EntityFrameworkCore;
using IndiaEventsWebApi;


var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddSwaggerGen();
builder.Services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
builder.Services.AddSingleton(new SemaphoreSlim(1, 1));
builder.Services.AddMemoryCache();
builder.Services.AddLazyCache();

var configuration = builder.Configuration;
//builder.Services.AddMySqlDataSource(builder.Configuration.GetConnectionString("SampleConnectionString")!);

builder.Services.AddDbContext<DataContext>(
    options =>
    {
        options.UseMySql(builder.Configuration.GetConnectionString("mysql"),
            new MySqlServerVersion(new Version(8, 0, 23)));
        //Microsoft.EntityFrameworkCore.ServerVersion.Parse("8.0.23-mysql"));
    });
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.SaveToken = true;
        options.RequireHttpsMetadata = false;
        options.TokenValidationParameters = new TokenValidationParameters
        {
            ValidateIssuer = true,
            ValidateAudience = true,
            ValidAudience = "ABM",
            ValidIssuer = "http://localhost:5098",
            IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes("veryveryveryveryverysecret......................"))
        };
    })
    .AddGoogle(options =>
    {
        options.ClientId = configuration["GoogleAuthentication:ClientId"];
        options.ClientSecret = configuration["GoogleAuthentication:ClientSecret"];
    });


builder.Services.Configure<IISServerOptions>(options =>
{
    options.MaxRequestBodySize = 1073741824; // 1 GB in bytes
    options.MaxRequestBodyBufferSize = 1073741824;
});

builder.Services.Configure<KestrelServerOptions>(options =>
{
    options.Limits.MaxRequestBodySize = 1073741824; // 1 GB in bytes
    options.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(20); //  20 minutes
});

builder.Services.AddCors(options =>
{
    options.AddPolicy("UseCors", builder =>
    {
        builder.WithOrigins("https://ambitious-rock-0757ea510.4.azurestaticapps.net", "http://localhost:4200")
               .AllowAnyMethod()
               .AllowAnyHeader();
    });
});
Log.Logger = new LoggerConfiguration()
            .WriteTo.File("Logs\\logs.txt", rollingInterval: RollingInterval.Day).CreateLogger();

//Log.Logger = new LoggerConfiguration()
//    .WriteTo.Console()
//    .CreateLogger();

//builder.Host.UseSerilog((hostingContext, loggerConfiguration) =>
//{
//    loggerConfiguration.ReadFrom.Configuration(hostingContext.Configuration)
//        .Enrich.FromLogContext()
//        .WriteTo.Console();
//}, preserveStaticLogger: true);

//builder.Host.UseSerilog((hostingContext, loggerConfiguration) =>
//{
//    loggerConfiguration
//        .ReadFrom.Configuration(hostingContext.Configuration)
//        .Enrich.FromLogContext()
//        .WriteTo.Logger(lc => lc.Filter.ByIncludingOnly(evt => evt.Exception != null).WriteTo.Console());
//}, preserveStaticLogger: true);

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI();
app.UseHttpsRedirection();
app.UseCors("UseCors");
//app.UseSerilogRequestLogging();
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.Run();

