using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Sahadeva.Dossier.DocumentGenerator.OpenXml
{
    internal class PlaceholderHelper
    {
        /// <summary>
        /// Looks for placeholders matching [AF.*]
        /// </summary>
        private readonly Regex _placeholder = new Regex(@"\[AF\.[^\]]+\](?!.*\[\[AF\.[^\]]+\]\])", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private readonly Regex _placeholderWithDataSource = new Regex(@"\[AF\.(?:Value|MultilineValue|Table|Url|Section\.Start):[^\]]+\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Only searches for placeholders that contain a data source i.e TableName.
        /// Children of Tables, Sections etc are ignored
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <returns></returns>
        internal List<Text> GetPlaceholdersWithDataSource(WordprocessingDocument wordDoc)
        {
            FixPlaceholdersAcrossRuns(wordDoc);
            IsolatePlaceholders(wordDoc);

            return ExtractDataSourcePlaceholdersFromDocument(wordDoc);
        }

        /// <summary>
        /// Gets all the placeholders in the document template
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"></exception>
        internal List<Text> GetAllPlaceholders(WordprocessingDocument wordDoc)
        {
            var body = wordDoc.MainDocumentPart?.Document.Body ?? throw new ApplicationException("Invalid document");

            return body.Descendants<Text>()
                .Where(e => _placeholder.IsMatch(e.Text))
                .ToList();
        }

        /// <summary>
        /// In some instances placeholders may be surrounded by text e.g when a placeholder is used in a sentence.
        /// This makes it difficult to replace the entire text node with a value as we would overwrite other text as well.
        /// This method will ensure that each placeholder sits in its own text node so that when we override/replace the placeholder
        /// it should not result in any data loss.
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void IsolatePlaceholders(WordprocessingDocument wordDoc)
        {
            XDocument xDoc = wordDoc.MainDocumentPart.GetXDocument();

            var textElements = xDoc.Descendants(W.t)
                   .Where(e => _placeholder.IsMatch(e.Value));

            foreach (var textElement in textElements.ToList())
            {
                string text = textElement.Value;
                var matches = _placeholder.Matches(text);

                if (matches.Count > 0)
                {
                    var newElements = new List<XElement>();
                    int currentIndex = 0;

                    foreach (Match match in matches)
                    {
                        int placeholderStartIndex = match.Index;
                        int placeholderLength = match.Length;

                        // Add text before the placeholder
                        if (placeholderStartIndex > currentIndex)
                        {
                            string beforePlaceholder = text.Substring(currentIndex, placeholderStartIndex - currentIndex);
                            if (!string.IsNullOrEmpty(beforePlaceholder))
                            {
                                newElements.Add(new XElement(
                                    W.t,
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    beforePlaceholder)
                                );
                            }
                        }

                        // Add the placeholder itself
                        string placeholder = match.Value;
                        newElements.Add(new XElement(
                            W.t,
                            new XAttribute(XNamespace.Xml + "space", "preserve"), 
                            placeholder));

                        // Update currentIndex to after the placeholder
                        currentIndex = placeholderStartIndex + placeholderLength;
                    }

                    // Add text after the last placeholder
                    if (currentIndex < text.Length)
                    {
                        string afterPlaceholder = text.Substring(currentIndex);
                        if (!string.IsNullOrEmpty(afterPlaceholder))
                        {
                            newElements.Add(new XElement(
                                W.t,
                                new XAttribute(XNamespace.Xml + "space", "preserve"),
                                afterPlaceholder)
                            );
                        }
                    }

                    // Replace the old text element with the new elements
                    textElement.AddBeforeSelf(newElements);

                    // Remove the original text element
                    textElement.Remove();
                }
            }

            // Save changes to the document
            wordDoc.MainDocumentPart.PutXDocument();
        }

        /// <summary>
        /// Extracts all placeholders that define a Table data source
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"></exception>
        private List<Text> ExtractDataSourcePlaceholdersFromDocument(WordprocessingDocument wordDoc)
        {
            var body = wordDoc.MainDocumentPart?.Document.Body ?? throw new ApplicationException("Invalid document");

            return body.Descendants<Text>()
                .Where(e => _placeholderWithDataSource.IsMatch(e.Text))
                .ToList();
        }

        /// <summary>
        /// At times, OpenXml may split text across multiple runs. This method flattens such occurrences so that the placeholders
        /// can be identified and replaced more easily.
        /// </summary>
        /// <param name="stream"></param>
        private void FixPlaceholdersAcrossRuns(WordprocessingDocument wordDoc)
        {
            XDocument xDoc = wordDoc.MainDocumentPart.GetXDocument();

            var content = xDoc.Descendants(W.p);
            var count = RegexHelper.Replace(
                content,
                _placeholder,
                (match) => match.Value);

            // Save changes to the document
            wordDoc.MainDocumentPart.PutXDocument();
        }
    }
}
