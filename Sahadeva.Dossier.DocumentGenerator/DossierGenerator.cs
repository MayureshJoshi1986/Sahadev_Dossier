using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DAL;
using Sahadeva.Dossier.DocumentGenerator.Extensions;
using Sahadeva.Dossier.DocumentGenerator.IO;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using Sahadeva.Dossier.DocumentGenerator.Processing;
using Sahadeva.Dossier.Entities;
using System.Data;

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

                // TODO: We can optimise here once we know from where to get the data
                // Currently, we seem to have multiple DAL classes which seem quite similar to each other
                var data = LoadDataset(job);

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

        private DataSet LoadDataset(DossierJob job)
        {
            var dataset = new DataSet();
            var dal = new dossierDAL();
            
            var clientData = dal.Dossier_FetchClientData_DT2(job.CoverageDossierId);
            dataset.AddTableToDataSet(clientData, "ClientData");

            var overviewData = dal.CoverageDossier_OverView_Page(job.CoverageDossierId);
            dataset.AddTableToDataSet(overviewData, "Overview");

            return dataset;
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
