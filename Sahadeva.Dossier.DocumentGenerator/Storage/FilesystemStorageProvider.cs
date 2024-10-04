using Microsoft.Extensions.Options;
using Sahadeva.Dossier.DocumentGenerator.Configuration;

namespace Sahadeva.Dossier.DocumentGenerator.Storage
{
    internal class FilesystemStorageProvider : IStorageProvider
    {
        private readonly FilesystemStorageOptions _options;

        public FilesystemStorageProvider(IOptions<FilesystemStorageOptions> options)
        {
            _options = options.Value;
        }

        public Task<byte[]> GetFile(string fileName)
        {
            var filePath = GetTemplatePath(fileName);
            return File.ReadAllBytesAsync(filePath);
        }

        public async Task WriteFile(MemoryStream stream, string fileName)
        {
            var filePath = GetOutputPath(fileName);

            // Ensure the directory exists
            var directory = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory!);
            }

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 4096, useAsync: true))
            {
                stream.Position = 0; // Ensure the stream's position is at the start
                await stream.CopyToAsync(fileStream);
            }
        }

        private string GetTemplatePath(string fileName) => Path.Combine(_options.TemplatePath, fileName);

        private string GetOutputPath(string fileName) => Path.Combine(_options.OutputPath, fileName);

    }
}
