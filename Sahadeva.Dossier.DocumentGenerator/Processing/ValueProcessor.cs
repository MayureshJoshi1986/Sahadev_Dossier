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
        private string _tableName = string.Empty;

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

            _tableName = config[0];
            _columnName = config[1];
        }

        public override void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data)
        {
            Placeholder.Text = GetDataFromSource(data);
        }

        protected string GetDataFromSource(DataSet data)
        {
            var table = data.Tables[_tableName] ?? throw new ApplicationException($"Could not find table '{_tableName}'");

            if (table.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Expression}' for multiple possible values"); }

            if (!table.Columns.Contains(_columnName)) { throw new ApplicationException($"Could not find column '{_columnName}' in '{_tableName}"); }

            return data.Tables[_tableName]!.Rows[0][_columnName].ToString()!;
        }
    }
}
