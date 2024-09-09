using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using Sahadeva.Dossier.DocumentGenerator.Parsers;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal class TablePlaceholderFactory
    {
        private readonly PlaceholderParser _placeholderParser;
        private readonly IServiceProvider _serviceProvider;

        public TablePlaceholderFactory(PlaceholderParser placeholderParser, IServiceProvider serviceProvider)
        {
            _placeholderParser = placeholderParser;
            _serviceProvider = serviceProvider;
        }

        internal ITablePlaceholderProcessor CreateProcessor(Text placeholder)
        {
            var placeholderType = _placeholderParser.GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "Table.Value" => ActivatorUtilities.CreateInstance<TableValueProcessor>(_serviceProvider, placeholder),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}"),
            };
        }
    }
}
