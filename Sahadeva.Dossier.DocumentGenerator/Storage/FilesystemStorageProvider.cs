using Sahadeva.Dossier.DocumentGenerator.Configuration;

namespace Sahadeva.Dossier.DocumentGenerator.Storage
{
    internal class FilesystemStorageProvider : IStorageProvider
    {
        private readonly FilesystemStorageOptions _options;

        public FilesystemStorageProvider(FilesystemStorageOptions options)
        {
            _options = options;
        }

        public Task<byte[]> GetFile(string fileName)
        {
            var filePath = GetTemplatePath(fileName);
            return File.ReadAllBytesAsync(filePath);
        }

        public void WriteFile(MemoryStream stream, string fileName)
        {
            var filePath = GetOutputPath(fileName);

            // Ensure the directory exists
            var directory = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory!);
            }

            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                stream.WriteTo(fileStream);
            }
        }

        private string GetTemplatePath(string fileName) => Path.Combine(_options.TemplatePath, fileName);

        private string GetOutputPath(string fileName) => Path.Combine(_options.OutputPath, fileName);

    }
}
