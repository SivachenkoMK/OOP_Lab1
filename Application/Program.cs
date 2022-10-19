using System;
using System.IO;
using System.Windows.Forms;
using Application.Configs;
using Application.Interfaces;
using Application.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Application
{
    public static class Program
    {
        private static IConfiguration? _configuration;
        private const string EnvVariableName = "ASPNETCORE_ENVIRONMENT";

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.SetHighDpiMode(HighDpiMode.SystemAware);
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            
            var envName = Environment.GetEnvironmentVariable(EnvVariableName);
            var settingsFileName = !string.IsNullOrEmpty(envName)
                ? $"appsettings.{envName}.json"
                : $"appsettings.Development.json";


            _configuration =
                new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(settingsFileName).Build();

            
            var host = CreateHostBuilder().Build();

            System.Windows.Forms.Application.Run(host.Services.GetRequiredService<MyExcel>());
        }

        private static IHostBuilder CreateHostBuilder()
        {
            return Host.CreateDefaultBuilder()
                .ConfigureServices((_, services)=>{
                    services.AddTransient<ICellService, CellService>();
                    services.AddTransient<ITableService, TableService>();
                    services.AddTransient<MyExcel>();
                    services.Configure<ErrorMessages>(_configuration!.GetSection(nameof(ErrorMessages)));
                    services.Configure<FileManagementOptions>(_configuration.GetSection(nameof(FileManagementOptions)));
                    services.Configure<DefaultConfiguration>(_configuration.GetSection(nameof(DefaultConfiguration)));
                });
        }
    }
}
