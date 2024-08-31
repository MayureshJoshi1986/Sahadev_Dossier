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
        public string TableName { get; private set; } = string.Empty;

        public string ColumnName { get; private set; } = string.Empty;

        public ValueProcessor(Text placeholder) : base(placeholder)
        {
        }

        protected override Regex CreatePlaceholderRegex()
        {
            return new Regex(@"(?<=\[AF\.Value:).*(?=\])", RegexOptions.Compiled);
        }

        protected override void ExtractPlaceholderParams()
        {
            var matches = PlaceholderRegex.Matches(Expression);

            if (matches.Count != 1)
            {
                throw new ApplicationException("Invalid expression for AF.Value. Required [AF.Value:<TableName>.<ColumnName>]");
            }

            var config = matches[0].Value.Split(".");

            TableName = config[0];
            ColumnName = config[1];
        }

        public override void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data)
        {
            Placeholder.Text = GetDataFromSource(data);
        }

        protected string GetDataFromSource(DataSet data)
        {
            var table = data.Tables[TableName] ?? throw new ApplicationException($"Could not find table '{TableName}'");

            if (table.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Expression}' for multiple possible values"); }

            if (!table.Columns.Contains(ColumnName)) { throw new ApplicationException($"Could not find column '{ColumnName}' in '{TableName}"); }

            return data.Tables[TableName]!.Rows[0][ColumnName].ToString()!;
        }
    }
}
