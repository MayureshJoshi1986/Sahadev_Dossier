using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using Sahadeva.Dossier.DocumentGenerator.Parsers;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal class PlaceholderFactory
    {
        private readonly PlaceholderParser _placeholderParser;
        private readonly IServiceProvider _serviceProvider;

        public PlaceholderFactory(PlaceholderParser placeholderParser, IServiceProvider serviceProvider)
        {
            _placeholderParser = placeholderParser;
            _serviceProvider = serviceProvider;
        }

        internal IPlaceholderWithDataSource CreateProcessor(Text placeholder, WordprocessingDocument document)
        {
            var placeholderType = _placeholderParser.GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "Value" => ActivatorUtilities.CreateInstance<DocumentValueProcessor>(_serviceProvider, placeholder),
                "MultilineValue" => ActivatorUtilities.CreateInstance<DocumentMultilineValueProcessor>(_serviceProvider, placeholder),
                "Table" => ActivatorUtilities.CreateInstance<TableProcessor>(_serviceProvider, placeholder, document),
                "Url" => ActivatorUtilities.CreateInstance<DocumentUrlProcessor>(_serviceProvider, placeholder, document),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}"),
            };
        }
    }
}
