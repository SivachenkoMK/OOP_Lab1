using Excel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace Excel
{
    public static class Program
    {
        private static IServiceProvider ServiceProvider { get; set; }

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            var host = CreateHostBuilder().Build();
            ServiceProvider = host.Services;

            Application.Run(ServiceProvider.GetRequiredService<MyExcel>());
        }

        private static IHostBuilder CreateHostBuilder()
        {
            return Host.CreateDefaultBuilder()
                .ConfigureServices((context, services)=>{
                    services.AddTransient<ICellService, CellService>();
                    services.AddTransient<ITableService, TableService>();
                    services.AddTransient<MyExcel>();
                });
        }

    }
}
