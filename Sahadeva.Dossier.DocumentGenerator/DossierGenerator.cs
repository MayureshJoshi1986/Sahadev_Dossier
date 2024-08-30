using DocumentFormat.OpenXml.Packaging;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class DossierGenerator
    {
        private readonly FileManager _fileManager;
        private readonly PlaceholderHelper _placeHolderHelper;

        public DossierGenerator(FileManager fileManager, PlaceholderHelper placeholderHelper)
        {
            _fileManager = fileManager;
            _placeHolderHelper = placeholderHelper;
        }

        internal async Task CreateDocumentFromTemplate(string templateName, string outputFileName)
        {
            using (MemoryStream stream = await ReadFromTemplate(templateName))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var placeholders = _placeHolderHelper.GetPlaceholderMap(wordDoc);

                // TODO: Temp to check placeholders
                placeholders.ForEach(x => Console.WriteLine(x.Text));

                // Flush changes from the word doc to the memory stream
                wordDoc.Save();

                WriteFile(stream, outputFileName);
            }
        }

        private async Task<MemoryStream> ReadFromTemplate(string fileName)
        {
            var content = await _fileManager.GetTemplate(fileName);
            return new MemoryStream(content);
        }

        private void WriteFile(MemoryStream stream, string fileName)
        {
            _fileManager.WriteFile(stream, fileName);
        }
    }
}
