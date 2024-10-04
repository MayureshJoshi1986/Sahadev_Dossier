
namespace Sahadeva.Dossier.DocumentGenerator.Storage
{
    internal interface IStorageProvider
    {
        Task<byte[]> GetFile(string fileName);

        Task WriteFile(MemoryStream stream, string fileName);
    }
}
