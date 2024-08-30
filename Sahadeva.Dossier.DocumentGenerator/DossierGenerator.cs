using DocumentFormat.OpenXml.Packaging;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using Sahadeva.Dossier.DocumentGenerator.Processing;
using Sahadeva.Dossier.Entities;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class DossierGenerator
    {
        private readonly FileManager _fileManager;
        private readonly PlaceholderHelper _placeholderHelper;
        private readonly PlaceholderFactory _placeholderFactory;

        public DossierGenerator(
            FileManager fileManager, 
            PlaceholderHelper placeholderHelper, 
            PlaceholderFactory placeholderFactory)
        {
            _fileManager = fileManager;
            _placeholderHelper = placeholderHelper;
            _placeholderFactory = placeholderFactory;
        }

        internal async Task ExecuteJob(DossierJob job)
        {
            using (MemoryStream stream = await ReadFromTemplate(job.TemplateName))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var placeholders = _placeholderHelper.GetPlaceholderMap(wordDoc);

                foreach (var placeholder in placeholders) 
                {
                    var processor = _placeholderFactory.CreateProcessor(placeholder);
                    //processor.ReplacePlaceholder(wordDoc);
                }

                // TODO: Temp to check placeholders
                placeholders.ForEach(x => Console.WriteLine(x.Text));

                // Flush changes from the word doc to the memory stream
                wordDoc.Save();

                WriteFile(stream, job.TemplateName);
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
