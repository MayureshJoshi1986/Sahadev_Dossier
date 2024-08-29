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
            var filePath = GetFinalPath(fileName);

            return File.ReadAllBytesAsync(filePath);
        }

        public void WriteFile(MemoryStream stream, string fileName)
        {
            var filePath = GetFinalPath(fileName);
            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                stream.WriteTo(fileStream);
            }
        }

        private string GetFinalPath(string fileName) => Path.Combine(_options.TemplatePath, fileName);
    }
}
