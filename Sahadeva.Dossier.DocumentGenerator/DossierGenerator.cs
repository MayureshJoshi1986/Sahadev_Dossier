using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Data;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using Sahadeva.Dossier.DocumentGenerator.Processors;
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
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
            {
                var placeholders = _placeholderHelper.GetPlaceholdersWithDataSource(document);

                var data = _datasetLoader.LoadDataset(job, placeholders);

                foreach (var placeholder in placeholders)
                {
                    var processor = _placeholderFactory.CreateProcessor(placeholder, document);
                    
                    Console.Write(placeholder.Text + "="); // TODO: For testing
                    
                    var dataTable = data.Tables[processor.TableName]
                        ?? throw new ApplicationException($"Could not find table for {placeholder.Text} having name {processor.TableName}");
                    processor.ReplacePlaceholder(dataTable);

                    Console.WriteLine(placeholder.Text); // TODO: For testing
                }

                CheckForUnProcessedPlaceholders(document);

                RemoveGrammarErrors(document);

                // Flush changes from the word doc to the memory stream
                document.Save();

                WriteFile(stream, job.TemplateName);
            }
        }

        /// <summary>
        /// Verify that we do not have any unprocessed placdeholders in the document
        /// </summary>
        /// <param name="document"></param>
        /// <exception cref="ApplicationException"></exception>
        private void CheckForUnProcessedPlaceholders(WordprocessingDocument document)
        {
            var leftOvers = _placeholderHelper.GetAllPlaceholders(document);

            if (leftOvers.Any())
            {
                throw new ApplicationException($"The document contains {leftOvers.Count} unprocessed placeholder(s)");
            }
        }

        /// <summary>
        /// Removes any grammar error marks in the document.
        /// This does not affect the document layout
        /// </summary>
        /// <param name="document"></param>
        private void RemoveGrammarErrors(WordprocessingDocument document)
        {
            var proofErrors = document.MainDocumentPart!.Document.Descendants<ProofError>().ToList();

            foreach (var error in proofErrors)
            {
                error.Remove();
            }
        }

        private async Task<MemoryStream> ReadFromTemplate(string fileName)
        {
            var content = await _fileManager.GetTemplate(fileName);
            var stream = new MemoryStream();
            stream.Write(content);

            return stream;
        }

        private void WriteFile(MemoryStream stream, string fileName)
        {
            _fileManager.WriteFile(stream, fileName);
        }
    }
}
