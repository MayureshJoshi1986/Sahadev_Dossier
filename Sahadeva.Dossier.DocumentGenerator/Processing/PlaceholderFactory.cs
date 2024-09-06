using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class PlaceholderFactory : PlaceholderFactoryBase<IPlaceholderWithDataSource>
    {
        private readonly FormatterParser _formatterParser;
        private readonly TablePlaceholderFactory _tablePlaceholderFactory;

        public PlaceholderFactory(FormatterParser formatterParser, TablePlaceholderFactory tablePlaceholderFactory)
        {
            _formatterParser = formatterParser;
            _tablePlaceholderFactory = tablePlaceholderFactory;
        }

        internal override IPlaceholderWithDataSource CreateProcessor(Text placeholder)
        {
            var placeholderType = GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "Value" => new ValueProcessor(placeholder, _formatterParser),
                "MultilineValue" => new MultilineValueProcessor(placeholder),
                "Table" => new TableProcessor(placeholder, _tablePlaceholderFactory),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}"),
            };
        }
    }
}
