using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sahadeva.Dossier.DocumentGenerator.Configuration;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.Storage;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class Program
    {
        static IHost _appHost = InitialiseHost();

        static async Task Main(string[] args)
        {
            InitialiseHost();

            var dossierGenerator = _appHost.Services.GetRequiredService<DossierGenerator>();

            // we want this to keep running and processing jobs as they become available
            //while (true)
            //{
                // TODO: Read jobs from the queue

                // TODO: Pass in the params required to generate the dossier. e.g. Template name
                await dossierGenerator.CreateDocumentFromTemplate("Air_India_Express.docx", "Air_India_Dossier.docx");
            //}
        }


        /// <summary>
        /// Bootstraps the application
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ApplicationException"></exception>
        static IHost InitialiseHost()
        {
            return Host.CreateDefaultBuilder()
            .ConfigureAppConfiguration((context, config) =>
            {
                config
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("NETCORE_ENVIRONMENT")}.json", optional: true)
                .AddEnvironmentVariables();
            })
            .ConfigureServices(GetApplicationServices)
            .Build();
        }

        static void GetApplicationServices(HostBuilderContext context, IServiceCollection services)
        {
            var storageProvider = context.Configuration.GetRequiredSection("Storage").GetValue<StorageProvider>("Provider");

            if (storageProvider == StorageProvider.Filesystem)
            {
                var options = context.Configuration.GetRequiredSection("Storage:Options").Get<FilesystemStorageOptions>()
                    ?? throw new ApplicationException("Filesystem options are missing from config");

                services.AddSingleton<IStorageProvider, FilesystemStorageProvider>(sp => new FilesystemStorageProvider(options));
            }
            else
            {
                throw new ApplicationException($"Storage provider could not be found. Please check config 'Storage:Provider'");
            }

            services.AddSingleton<FileManager>();
            services.AddSingleton<DossierGenerator>();
        }
    }
}
