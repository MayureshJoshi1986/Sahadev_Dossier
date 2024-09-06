using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class RowValueProcessor : PlaceholderProcessorBase<DataRow>
    {
        protected string ColumnName { get; private set; } = string.Empty;

        private readonly FormatterParser _formatterParser;

        public RowValueProcessor(Text placeholder, FormatterParser formatterParser) : base(placeholder)
        {
            _formatterParser = formatterParser;
        }

        public override void ParsePlaceholder()
        {
            var match = Regex.Match(Placeholder.Text, @"(?<=\[AF\.RowValue:)[^;\|\]]+");

            if (match.Success)
            {
                ColumnName = match.Value;
            }
            else
            {
                throw new ApplicationException($"Could not parse {Placeholder.Text}");
            }
        }

        public override void ReplacePlaceholder(DataRow row)
        {
            if (string.IsNullOrWhiteSpace(ColumnName) || !row.Table.Columns.Contains(ColumnName)) { throw new ApplicationException($"Column name missing or invalid: '{ColumnName}'"); }
            var value = row[ColumnName].ToString()!;
            var formatter = _formatterParser?.GetFormatter(Placeholder.Text);
            Placeholder.Text = formatter?.Format(value) ?? value;
        }
    }
}
