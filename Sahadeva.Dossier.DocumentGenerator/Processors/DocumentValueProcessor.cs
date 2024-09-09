using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Formatters;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal partial class DocumentValueProcessor : DocumentPlaceholderProcessorBase, IPlaceholderWithDataSource
    {
        private readonly FormatterFactory? _formatterFactory;

        protected string ColumnName { get; private set; } = string.Empty;

        public DocumentValueProcessor(Text placeholder) : base(placeholder)
        {
            _formatterFactory = null;
        }

        public DocumentValueProcessor(Text placeholder, FormatterFactory formatterFactory) : base(placeholder)
        {
            _formatterFactory = formatterFactory;
        }

        public override void SetPlaceholderOptions()
        {
            var match = GetPlaceholderOptionsRegex().Match(Placeholder.Text);
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

        protected virtual Regex GetPlaceholderOptionsRegex()
        {
            return OptionsRegex();
        }

        public override void ReplacePlaceholder(DataTable data)
        {
            var value = GetValueFromSource(data);
            var formatter = _formatterFactory?.CreateFormatter(Placeholder);
            Placeholder.Text = formatter?.Format(value) ?? value;
        }

        protected string GetValueFromSource(DataTable data)
        {
            if (data.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Placeholder.Text}' for multiple possible values"); }

            if (!data.Columns.Contains(ColumnName)) { throw new ApplicationException($"Could not find column '{ColumnName}' in '{TableName}"); }

            return data.Rows[0][ColumnName].ToString()!;
        }

        [GeneratedRegex(@"\[AF\.Value:(?<TableName>[^\.\]]+)\.(?<ColumnName>[^\|\]]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex OptionsRegex();
    }
}
