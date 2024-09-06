using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class TableProcessor : PlaceholderProcessorBase<DataTable>, IPlaceholderWithDataSource
    {
        public string TableName { get; private set; } = string.Empty;

        private readonly Regex _placeholderDataSourceRegex = new Regex(@"(?<=\[AF\.Table:)[^\]]+", RegexOptions.Compiled);
        
        private readonly TablePlaceholderFactory _tablePlaceholderFactory;

        public TableProcessor(Text placeholder, TablePlaceholderFactory tablePlaceholderFactory) : base(placeholder)
        {
            _tablePlaceholderFactory = tablePlaceholderFactory;
        }

        public override void ParsePlaceholder()
        {
            var match = _placeholderDataSourceRegex.Match(Placeholder.Text);
            if (match.Success)
            {
                TableName = match.Value;
            }
            else
            {
                throw new ApplicationException($"Could not parse {Placeholder.Text}");
            }
        }

        public override void ReplacePlaceholder(DataTable data)
        {
            // Ensure that the Table placeholder is correctly placed within a table
            var table = Placeholder.Ancestors<Table>().FirstOrDefault()
                ?? throw new ApplicationException($"{Placeholder.Text} is only valid inside a table.");

            // Currently throwing if we have no data for the table but we could choose to hide he table as well, or define some placeholder text
            if (data.Rows.Count == 0) { throw new ApplicationException($"No data for {Placeholder.Text}"); }

            // The first row of the table MUST contain a Table placeholder which defines the datasource
            var tableNameRow = table.Elements<TableRow>().First();
            var isValidTableNameRow = tableNameRow.Descendants<Text>().Any(t => _placeholderDataSourceRegex.IsMatch(t.Text));
            if (!isValidTableNameRow) { throw new ApplicationException("The first row of the table MUST contain the table placeholder which defines the datasource"); }

            // Delete the first row (row containing the table name) as it is only required for processing and should not appear in the output
            tableNameRow.Remove();

            var templateRows = table.Elements<TableRow>().ToList();

            foreach (var row in templateRows)
            {
                ProcessRow(row, data);
            }
        }

        private void ProcessRow(TableRow row, DataTable dataTable)
        {
            var placeholders = ExtractRowPlaceholders(row);

            // No placeholders found in this row, nothing to process
            if (placeholders.Count == 0) { return; }

            // Check if any placeholder in the row has a filter and apply it, ideally it would be the first placeholder in the row
            string filterCriteria = string.Empty;
            foreach (var placeholder in placeholders)
            {
                var filter = ExtractFilterCriteria(placeholder);
                if (!string.IsNullOrWhiteSpace(filter))
                {
                    filterCriteria = filter;
                    break;
                }
            }

            DataRow[] filteredRows = filterCriteria.Length > 0 ? dataTable.Select(filterCriteria) : dataTable.Select();

            if (filteredRows.Length == 0)
            {
                // TODO: Should we throw here?
                Console.WriteLine($"No data found for {placeholders.First()}");
            }

            foreach (var dataRow in filteredRows)
            {
                // Create a clone of the template row so that we can replace the placeholders
                var clone = (TableRow)row.CloneNode(true);

                var textNodes = clone.Descendants<Text>();
                foreach (var textNode in textNodes)
                {
                    // Check if this is a placeholder node
                    if (placeholders.TryGetValue(textNode.Text, out var placeholder))
                    {
                        var tablePlaceholder = _tablePlaceholderFactory.CreateProcessor(textNode);
                        tablePlaceholder.ReplacePlaceholder(dataRow);
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

            var textElements = row.Descendants<Text>();
            foreach (var textElement in textElements)
            {
                var placeholderMatch = Regex.Match(textElement.Text, @"\[AF\.RowValue:[^\]]+\]");
                if (placeholderMatch.Success)
                {
                    placeholders.Add(placeholderMatch.Value);
                }
            }

            return placeholders;
        }

        private static string ExtractFilterCriteria(string placeholder)
        {
            var match = Regex.Match(placeholder, @"\[AF\.RowValue:(?<ColumnName>[^\;]+);Filter=(?<FilterColumn>[^\(]+)\((['‘’](?<FilterValue>[^'‘’]+)['‘’])\)\]");

            if (match.Success)
            {
                var columnName = match.Groups["FilterColumn"].Value;
                var value = match.Groups["FilterValue"].Value;
                return $"{columnName} = '{value}'";
            }

            // Return empty filter if no match
            return string.Empty;
        }
    }
}
