using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Data;
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
        private readonly DatasetLoader _datasetLoader;

        public DossierGenerator(
            FileManager fileManager, 
            PlaceholderHelper placeholderHelper, 
            PlaceholderFactory placeholderFactory,
            DatasetLoader datasetLoader)
        {
            _fileManager = fileManager;
            _placeholderHelper = placeholderHelper;
            _placeholderFactory = placeholderFactory;
            _datasetLoader = datasetLoader;
        }

        internal async Task ExecuteJob(DossierJob job)
        {
            using (MemoryStream stream = await ReadFromTemplate(job.TemplateName))
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var placeholders = _placeholderHelper.GetPlaceholderMap(wordDoc);

                var data = _datasetLoader.LoadDataset(job, placeholders);

                foreach (var placeholder in placeholders) 
                {
                    var processor = _placeholderFactory.CreateProcessor(placeholder);
                    Console.Write(placeholder.Text + "="); // TODO: For testing
                    processor.ReplacePlaceholder(wordDoc, data);
                    Console.WriteLine(placeholder.Text); // TODO: For testing
                }

                RemoveGrammarErrors(wordDoc);

                // Flush changes from the word doc to the memory stream
                wordDoc.Save();

                WriteFile(stream, job.TemplateName);
            }
        }

        /// <summary>
        /// Removes any grammar error marks in the document.
        /// This does not affect the document layout
        /// </summary>
        /// <param name="wordDoc"></param>
        public void RemoveGrammarErrors(WordprocessingDocument wordDoc)
        {
            var proofErrors = wordDoc.MainDocumentPart!.Document.Descendants<ProofError>().ToList();

            foreach (var error in proofErrors)
            {
                error.Remove();
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
