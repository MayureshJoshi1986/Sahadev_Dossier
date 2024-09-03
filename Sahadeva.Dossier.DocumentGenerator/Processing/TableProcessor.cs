using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class TableProcessor : PlaceholderProcessorBase
    {
        private string _tableName = string.Empty;

        public TableProcessor(Text placeholder) : base(placeholder)
        {
        }

        public override void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data)
        {
            //throw new NotImplementedException();
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

            _tableName = matches[0].Value;
        }
    }
}
