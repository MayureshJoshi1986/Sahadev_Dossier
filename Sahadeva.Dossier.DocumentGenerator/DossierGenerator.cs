using DocumentFormat.OpenXml.Packaging;
using Sahadeva.Dossier.DocumentGenerator.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using DocumentFormat.OpenXml;
using OpenXmlPowerTools;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;

namespace Sahadeva.Dossier.DocumentGenerator
{
    internal class DossierGenerator
    {
        private readonly Regex _placeholder = new Regex(@"\{\{AF\.[^\}]+\}\}");
        private readonly FileManager _fileManager;

        public DossierGenerator(FileManager fileManager)
        {
            _fileManager = fileManager;
        }

        internal async Task CreateDocumentFromTemplate(string templateName, string outputFileName)
        {
            using (MemoryStream stream = await ReadFromTemplate(templateName))
            {
                FixPlaceholdersAcrossRuns(stream);
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

        private void FixPlaceholdersAcrossRuns(MemoryStream stream)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                XDocument xDoc = wordDoc.MainDocumentPart.GetXDocument();

                var content = xDoc.Descendants(W.p);
                var count = RegexHelper.Replace(
                    content,
                    _placeholder,
                    (match) => match.Value);

                // Save changes to the document
                wordDoc.MainDocumentPart.PutXDocument();

                //var body = wordDoc.MainDocumentPart?.Document.Body ?? throw new ApplicationException("Invalid document");
                // Replace placeholders in the document body
                //foreach (var text in body.Descendants<Text>())
                //{
                //    foreach (var placeholder in replacements.Keys)
                //    {
                //        if (text.Text.Contains(placeholder))
                //        {
                //            text.Text = text.Text.Replace(placeholder, replacements[placeholder]);
                //        }
                //    }
                //}
                //wordDoc.MainDocumentPart.Document.Save();
            }

        }
    }
}
