
namespace Sahadeva.Dossier.DocumentGenerator.Storage
{
    internal interface IStorageProvider
    {
        Task<byte[]> GetFile(string fileName);

        void WriteFile(MemoryStream stream, string fileName);
    }
}
