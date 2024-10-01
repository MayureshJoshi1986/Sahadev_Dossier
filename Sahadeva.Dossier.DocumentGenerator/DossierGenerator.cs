using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Sahadeva.Dossier.DocumentGenerator.Data;
using Sahadeva.Dossier.DocumentGenerator.Imaging;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using Sahadeva.Dossier.DocumentGenerator.Processors;
using Sahadeva.Dossier.Entities;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class DossierGenerator
    {
        private readonly DocumentHelper _documentHelper;
        private readonly FileManager _fileManager;
        private readonly PlaceholderHelper _placeholderHelper;
        private readonly PlaceholderFactory _placeholderFactory;
        private readonly DatasetLoader _datasetLoader;
        private readonly ImageDownloader _imageDownloader;

        public DossierGenerator(
            DocumentHelper documentHelper,
            FileManager fileManager,
            PlaceholderHelper placeholderHelper,
            PlaceholderFactory placeholderFactory,
            DatasetLoader datasetLoader,
            ImageDownloader imageDownloader)
        {
            _documentHelper = documentHelper;
            _fileManager = fileManager;
            _placeholderHelper = placeholderHelper;
            _placeholderFactory = placeholderFactory;
            _datasetLoader = datasetLoader;
            _imageDownloader = imageDownloader;
        }

        internal async Task ExecuteJob(DossierJob job)
        {
            using (MemoryStream stream = await ReadFromTemplate(job.TemplateName))
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, true))
            {
                _documentHelper.StripTrackingInfo(document);

                _placeholderHelper.FixPlaceholdersAcrossRuns(document);
                _placeholderHelper.IsolatePlaceholders(document);

                var placeholders = _placeholderHelper.GetPlaceholdersWithDataSource(document);

                var data = _datasetLoader.LoadDataset(job, placeholders.Select(p => p.Text));

                foreach (var placeholder in placeholders)
                {
                    var processor = _placeholderFactory.CreateProcessor(job, placeholder, document);

                    var dataTable = data.Tables[processor.TableName]
                        ?? throw new ApplicationException($"Could not find table for {placeholder.Text} having name {processor.TableName}");
                    
                    processor.ReplacePlaceholder(dataTable);
                }

                await _imageDownloader.DownloadImagesAsync(document);

                CheckForUnProcessedPlaceholders(document);

                _documentHelper.RemoveGrammarErrors(document);

                // TODO: Check the template for the errors so we know if the issues are after generation or existing
                //OpenXmlValidator validator = new OpenXmlValidator();
                //int errorCount = 0;

                //foreach (ValidationErrorInfo error in validator.Validate(document))
                //{
                //    Console.WriteLine("Error Description: {0}", error.Description);
                //    Console.WriteLine("Error Path: {0}", error.Path.XPath);
                //    Console.WriteLine("Error Part: {0}", error.Part.Uri);
                //    errorCount++;
                //}

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
