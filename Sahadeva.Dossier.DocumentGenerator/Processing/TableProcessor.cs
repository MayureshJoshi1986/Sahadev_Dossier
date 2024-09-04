using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class TableProcessor : PlaceholderProcessorBase
    {
        public TableProcessor(Text placeholder) : base(placeholder)
        {
        }

        public override void ReplacePlaceholder(WordprocessingDocument wordDoc, DataTable data)
        {
            var table = Placeholder.Ancestors<Table>().FirstOrDefault() 
                ?? throw new ApplicationException($"{Placeholder.Text} is only valid inside a table.");
            
            if (data.Rows.Count == 0) { throw new ApplicationException($"No data for {Placeholder.Text}"); }

            // The first row of the table MUST be the table placeholder which defines the datasource
            var tableNameRow = table.Elements<TableRow>().First();
            var isValidTableNameRow = tableNameRow.Descendants<Text>().Any(t => PlaceholderOptionsRegex.IsMatch(t.Text));

            if (!isValidTableNameRow) { throw new ApplicationException("The first row of the table MUST contain the table placeholder which defines the datasource"); }

            // Delete the table name row as it is only required for processing and should not appear in the output
            tableNameRow.Remove();

            var templateRows = table.Elements<TableRow>().ToList();

            foreach (var row in templateRows)
            {
                ProcessRow(row, data);
            }
        }

        protected override Regex GetPlaceholderOptionsRegex()
        {
            return new Regex(@"(?<=\[AF\.Table:)[^\]]+", RegexOptions.Compiled);
        }

        protected override void ExtractPlaceholderOptions()
        {
            var matches = PlaceholderOptionsRegex.Matches(Expression);

            if (matches.Count != 1)
            {
                throw new ApplicationException("Invalid expression for AF.Table. Required [AF.Table:<TableName>]");
            }

            DataSourceName = matches[0].Value;
        }

        private static void ProcessRow(TableRow row, DataTable dataTable)
        {
            var placeholders = ExtractRowPlaceholders(row);

            // No placeholders found, nothing to process
            if (placeholders.Count == 0) { return; }

            // Check if the first placeholder in the row has a filter and apply it
            var filterCriteria = ExtractFilterCriteria(placeholders.First());
            DataRow[] filteredRows = filterCriteria.Length > 0 ? dataTable.Select(filterCriteria) : dataTable.Select();

            if (filteredRows.Length == 0)
            {
                Console.WriteLine($"No data found for {placeholders.First()}");
            }

            foreach (var dataRow in filteredRows)
            {
                // Create a clone of the template row so that we can replace the placeholders
                var clone = (TableRow)row.CloneNode(true);

                var textNodes = clone.Descendants<Text>();
                foreach (var textNode in textNodes)
                {
                    // Check if text matches a placeholder and replace it
                    if (placeholders.TryGetValue(textNode.Text, out var placeholder))
                    {
                        ReplacePlaceholder(textNode, placeholder, dataRow);
                    }
                }

                row.InsertBeforeSelf(clone);
            }

            // Remove the original row as we have cloned the required rows
            row.Remove();
        }

        private static HashSet<string> ExtractRowPlaceholders(TableRow row)
        {
            var placeholders = new HashSet<string>();

            foreach (var cell in row.Elements<TableCell>())
            {
                // Here you would extract the placeholder from the text inside the cell
                var textElements = cell.Descendants<Text>();
                foreach (var textElement in textElements)
                {
                    var placeholderMatch = Regex.Match(textElement.Text, @"\[AF\.TableRow:[^\]]+\]");
                    if (placeholderMatch.Success)
                    {
                        placeholders.Add(placeholderMatch.Value);
                    }
                }
            }

            return placeholders;
        }

        private static string ExtractFilterCriteria(string placeholder)
        {
            var match = Regex.Match(placeholder, @"\[AF\.TableRow:(?<ColumnName>[^\;]+);Filter=(?<FilterColumn>[^\(]+)\((['‘’](?<FilterValue>[^'‘’]+)['‘’])\)\]");

            if (match.Success)
            {
                var columnName = match.Groups["FilterColumn"].Value;
                var value = match.Groups["FilterValue"].Value;
                return $"{columnName} = '{value}'";
            }

            // Return empty filter if no match
            return string.Empty;
        }

        private static void ReplacePlaceholder(Text textElement, string placeholder, DataRow dataRow)
        {
            var columnName = ExtractColumnName(placeholder);
            if (string.IsNullOrWhiteSpace(columnName) || !dataRow.Table.Columns.Contains(columnName)) { throw new ApplicationException($"Column name missing or invalid: '{columnName}'"); }

            textElement.Text = dataRow[columnName].ToString()!;
        }

        private static string ExtractColumnName(string placeholder)
        {
            var match = Regex.Match(placeholder, @"(?<=\[AF\.TableRow:)[^\];]+");

            return match.Success ? match.Value : string.Empty;
        }

    }
}
