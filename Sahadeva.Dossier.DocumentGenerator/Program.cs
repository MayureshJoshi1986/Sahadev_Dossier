using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Sahadeva.Dossier.DAL;
using Sahadeva.Dossier.DocumentGenerator.Configuration;
using Sahadeva.Dossier.DocumentGenerator.Data;
using Sahadeva.Dossier.DocumentGenerator.Formatters;
using Sahadeva.Dossier.DocumentGenerator.Imaging;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using Sahadeva.Dossier.DocumentGenerator.Parsers;
using Sahadeva.Dossier.DocumentGenerator.Processors;
using Sahadeva.Dossier.DocumentGenerator.Storage;
using ConfigurationManager = Sahadeva.Dossier.Common.Configuration.ConfigurationManager;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class Program
    {
        static IHost _appHost = InitialiseHost();

        static async Task Main(string[] args)
        {
            InitialiseHost();


            // we want this to keep running and processing jobs as they become available
            //while (true)
            //{

            try
            {
                // TODO: Read jobs from the queue
                // TODO: This is temp to simulate getting a job from the queue
                var job = DossierJobGenerator.GetJob(args);

                var dossierGenerator = _appHost.Services.GetRequiredService<DossierGenerator>();
                await dossierGenerator.ExecuteJob(job);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

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
            .ConfigureServices(GetApplicationServices)
            .Build();
        }

        static void GetApplicationServices(HostBuilderContext context, IServiceCollection services)
        {
            var configuration = ConfigurationManager.Settings;
            var storageProvider = configuration.GetRequiredSection("Storage").GetValue<StorageProvider>("Provider");

            if (storageProvider == StorageProvider.Filesystem)
            {
                var options = configuration.GetRequiredSection("Storage:Options").Get<FilesystemStorageOptions>()
                    ?? throw new ApplicationException("Filesystem options are missing from config");

                services.AddSingleton<IStorageProvider, FilesystemStorageProvider>(sp => new FilesystemStorageProvider(options));
            }
            else
            {
                throw new ApplicationException($"Storage provider could not be found. Please check config 'Storage:Provider'");
            }

            services.AddSingleton<FileManager>();
            services.AddSingleton<DocumentHelper>();
            services.AddSingleton<PlaceholderHelper>();
            services.AddSingleton<PlaceholderParser>();
            services.AddSingleton<PlaceholderFactory>();
            services.AddSingleton<FormatterFactory>();
            services.AddSingleton<RowPlaceholderFactory>();
            services.AddSingleton<DatasetLoader>();
            services.AddSingleton<ScreenshotService>();
            services.AddSingleton<DossierDAL>();
            services.AddSingleton<ImageDownloader>();

            services.AddTransient<DossierGenerator>();

            services.AddOptions<ScreenshotOptions>().Bind(configuration.GetSection(ScreenshotOptions.ConfigKey));
        }
    }
}
