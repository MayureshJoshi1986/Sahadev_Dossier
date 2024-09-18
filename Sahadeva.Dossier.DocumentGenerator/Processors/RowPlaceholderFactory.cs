using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using Sahadeva.Dossier.DocumentGenerator.Parsers;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal class RowPlaceholderFactory
    {
        private readonly PlaceholderParser _placeholderParser;
        private readonly IServiceProvider _serviceProvider;

        public RowPlaceholderFactory(PlaceholderParser placeholderParser, IServiceProvider serviceProvider)
        {
            _placeholderParser = placeholderParser;
            _serviceProvider = serviceProvider;
        }

        internal IRowPlaceholderProcessor CreateProcessor(Text placeholder, WordprocessingDocument document)
        {
            var placeholderType = _placeholderParser.GetPlaceholderType(placeholder.Text);

            return placeholderType switch
            {
                "Row.Value" => ActivatorUtilities.CreateInstance<RowValueProcessor>(_serviceProvider, placeholder),
                "Row.Url" => ActivatorUtilities.CreateInstance<RowUrlProcessor>(_serviceProvider, placeholder, document),
                _ => throw new NotSupportedException($"Unsupported placeholder type: {placeholderType} found in {placeholder.Text}"),
            };
        }
    }
}
