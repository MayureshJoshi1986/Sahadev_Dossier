using DocumentFormat.OpenXml.Packaging;
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

        internal ITablePlaceholderProcessor CreateProcessor(Text placeholder, WordprocessingDocument document)
        {
            var placeholderType = _placeholderParser.GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "Table.Value" => ActivatorUtilities.CreateInstance<TableValueProcessor>(_serviceProvider, placeholder),
                "Table.Url" => ActivatorUtilities.CreateInstance<TableUrlProcessor>(_serviceProvider, placeholder, document),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}"),
            };
        }
    }
}
