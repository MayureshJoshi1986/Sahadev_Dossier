using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Parsers;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal partial class SectionProcessor(
        Text placeholder,
        WordprocessingDocument document,
        PlaceholderParser placeholderParser,
        RowPlaceholderFactory rowPlaceholderFactory
            ) : PlaceholderProcessorBase(placeholder), IPlaceholderWithDataSource
    {
        private readonly WordprocessingDocument _document = document;
        private readonly PlaceholderParser _placeholderParser = placeholderParser;
        private readonly RowPlaceholderFactory _rowPlaceholderFactory = rowPlaceholderFactory;

        public string TableName { get; private set; } = string.Empty;

        public void ReplacePlaceholder(DataTable data)
        {
            var parentParagraph = GetParentParagraph();

            if (parentParagraph == null) { throw new ApplicationException($"Section placeholders should be placed in its own paragraph. {Placeholder.Text}"); }

            var sectionTemplate = GetSectionContent();
            var placeholders = ExtractSectionPlaceholders(sectionTemplate);

            var filter = _placeholderParser.GetFilter(Placeholder.Text);
            DataRow[] filteredRows = filter.Length > 0 ? data.Select(filter) : data.Select();

            //TODO: Remove Take 2
            foreach (var dataRow in filteredRows.Take(2))
            {
                // Create a clone of all the elements in the section template
                var clones = sectionTemplate.Select(n => n.CloneNode(true)).ToList();

                // Get all text nodes in the cloned section
                var textNodes = clones.SelectMany(c => c.Descendants<Text>()).ToList();
                foreach (var textNode in textNodes)
                {
                    // Check if this is a placeholder node
                    if (placeholders.TryGetValue(textNode.Text, out var placeholder))
                    {
                        var tablePlaceholder = _rowPlaceholderFactory.CreateProcessor(textNode, _document);
                        tablePlaceholder.ReplacePlaceholder(dataRow);
                    }
                }

                clones.ForEach(n => parentParagraph!.InsertBeforeSelf(n));
            }

            // Remove the original placeholder and placeholder content
            parentParagraph.Remove();
            sectionTemplate.ForEach(e => e.Remove());
        }

        private HashSet<string> ExtractSectionPlaceholders(List<OpenXmlElement> sectionTemplate)
        {
            var placeholders = new HashSet<string>();

            var textElements = sectionTemplate.SelectMany(e => e.Descendants<Text>());
            foreach (var textElement in textElements)
            {
                var placeholderMatch = SectionPlaceholderRegex().Match(textElement.Text);
                if (placeholderMatch.Success)
                {
                    placeholders.Add(placeholderMatch.Value);
                }
            }

            return placeholders;
        }

        public override void SetPlaceholderOptions()
        {
            var match = OptionsRegex().Match(Placeholder.Text);
            if (match.Success)
            {
                TableName = match.Value;
            }
            else
            {
                throw new ApplicationException($"Could not parse {Placeholder.Text}");
            }
        }

        private List<OpenXmlElement> GetSectionContent()
        {
            var insideSection = true;
            var sectionContent = new List<OpenXmlElement>();

            var parentParagraph = GetParentParagraph();
            var currentElement = parentParagraph?.NextSibling();

            while (currentElement != null && insideSection)
            {
                var end = currentElement
                    .Descendants<Text>()
                    .Where(t => SectionEndRegex().IsMatch(t.Text))
                    .FirstOrDefault();

                if (end != null)
                {
                    insideSection = false;
                    currentElement.Remove();
                    break;
                }

                sectionContent.Add(currentElement);

                currentElement = currentElement.NextSibling();
            }
            return sectionContent;
        }

        /// <summary>
        /// Section placeholders should sit alone in their own paragraphs. This is an important assumption in the processing of these placeholders
        /// </summary>
        /// <returns></returns>
        private Paragraph? GetParentParagraph()
        {
            OpenXmlElement? parentElement = Placeholder.Parent;

            // Traverse up the hierarchy until we find a Paragraph element
            while (parentElement != null && !(parentElement is Paragraph))
            {
                parentElement = parentElement.Parent;
            }

            return parentElement as Paragraph;
        }


        [GeneratedRegex(@"(?<=\[AF\.Section\.Start:)[^\]]+", RegexOptions.Compiled | RegexOptions.IgnoreCase)]
        private static partial Regex OptionsRegex();

        [GeneratedRegex(@"\[AF\.Section\.End\]", RegexOptions.Compiled | RegexOptions.IgnoreCase)]
        private static partial Regex SectionEndRegex();


        [GeneratedRegex(@"\[AF\.Row\.[^\]]+\]", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex SectionPlaceholderRegex();
    }
}
