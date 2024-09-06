using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    /// <summary>
    /// Replaces a placeholder with a single value
    /// </summary>
    internal class ValueProcessor : PlaceholderProcessorBase<DataTable>, IPlaceholderWithDataSource
    {
        private readonly FormatterParser? _formatterParser;

        public string TableName { get; private set; } = string.Empty;

        protected string ColumnName { get; private set; } = string.Empty;

        public ValueProcessor(Text placeholder) : base(placeholder)
        {

        }

        public ValueProcessor(Text placeholder, FormatterParser formatterParser) : base(placeholder)
        {
            _formatterParser = formatterParser;
        }

        public override void ParsePlaceholder()
        {
            var match = GetPlaceholderDataSourceRegex().Match(Placeholder.Text);
            if (match.Success)
            {
                TableName = match.Groups["TableName"].Value;
                ColumnName = match.Groups["ColumnName"].Value;
            }
            else
            {
                throw new ApplicationException($"Could not parse {Placeholder.Text}");
            }
        }

        public override void ReplacePlaceholder(DataTable data)
        {
            var value = GetDataFromSource(data);
            var formatter = _formatterParser?.GetFormatter(Placeholder.Text);
            Placeholder.Text = formatter?.Format(value) ?? value;
        }

        protected virtual Regex GetPlaceholderDataSourceRegex()
        {
            return new Regex(@"\[AF\.(Value):(?<TableName>[^\.\]]+)\.(?<ColumnName>[^\|\]]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        }

        protected string GetDataFromSource(DataTable data)
        {
            if (data.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Placeholder.Text}' for multiple possible values"); }

            if (!data.Columns.Contains(ColumnName)) { throw new ApplicationException($"Could not find column '{ColumnName}' in '{TableName}"); }

            return data.Rows[0][ColumnName].ToString()!;
        }
    }
}
