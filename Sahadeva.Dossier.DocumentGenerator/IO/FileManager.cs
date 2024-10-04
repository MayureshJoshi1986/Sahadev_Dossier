using Sahadeva.Dossier.DocumentGenerator.Storage;

namespace Sahadeva.Dossier.DocumentGenerator.IO
{
    internal class FileManager
    {
        private readonly IStorageProvider _storageProvider;

        public FileManager(IStorageProvider storageProvider)
        {
            _storageProvider = storageProvider;
        }

        internal Task<byte[]> GetTemplate(string fileName)
        {
            return _storageProvider.GetFile(fileName);
        }

        internal Task WriteFile(MemoryStream stream, string fileName)
        {
            return _storageProvider.WriteFile(stream, fileName);
        }
    }
}
