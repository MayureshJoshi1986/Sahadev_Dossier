using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    /// <summary>
    /// Replaces a placeholder with a single value
    /// </summary>
    internal class ValueProcessor : PlaceholderProcessorBase
    {
        private string _columnName = string.Empty;

        public ValueProcessor(Text placeholder) : base(placeholder)
        {
        }

        protected override Regex GetPlaceholderOptionsRegex()
        {
            return new Regex(@"(?<=\[AF\.Value:).*(?=\])", RegexOptions.Compiled);
        }

        protected override void ExtractPlaceholderOptions()
        {
            var matches = PlaceholderOptionsRegex.Matches(Expression);

            if (matches.Count != 1)
            {
                throw new ApplicationException("Invalid expression for AF.Value. Required [AF.Value:<TableName>.<ColumnName>]");
            }

            var config = matches[0].Value.Split(".");

            DataSourceName = config[0];
            _columnName = config[1];
        }

        public override void ReplacePlaceholder(WordprocessingDocument wordDoc, DataTable data)
        {
            Placeholder.Text = GetDataFromSource(data);
        }

        protected string GetDataFromSource(DataTable data)
        {
            if (data.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Expression}' for multiple possible values"); }

            if (!data.Columns.Contains(_columnName)) { throw new ApplicationException($"Could not find column '{_columnName}' in '{DataSourceName}"); }

            return data.Rows[0][_columnName].ToString()!;
        }
    }
}
