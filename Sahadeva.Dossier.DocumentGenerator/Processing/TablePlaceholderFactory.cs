using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class TablePlaceholderFactory : PlaceholderFactoryBase<IPlaceholderProcessor<DataRow>>
    {
        private readonly FormatterParser _formatterParser;

        public TablePlaceholderFactory(FormatterParser formatterParser)
        {
            _formatterParser = formatterParser;
        }

        internal override IPlaceholderProcessor<DataRow> CreateProcessor(Text placeholder)
        {
            var placeholderType = GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "RowValue" => new RowValueProcessor(placeholder, _formatterParser),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}"),
            };
        }
    }
}
