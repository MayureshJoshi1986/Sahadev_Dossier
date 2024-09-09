using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Parsers;
using Sahadeva.Dossier.DocumentGenerator.Formatters;

namespace Sahadeva.Dossier.DocumentGenerator.Formatters
{
    internal class FormatterFactory
    {
        private readonly PlaceholderParser _placeholderParser;

        public FormatterFactory(PlaceholderParser placeholderParser)
        {
            _placeholderParser = placeholderParser;
        }

        internal IValueFormatter CreateFormatter(Text placeholder)
        {
            var formatSpecifier = _placeholderParser.GetFormatter(placeholder.Text);

            if (string.IsNullOrWhiteSpace(formatSpecifier))
            {
                return new NoOpFormatter();
            }

            if (formatSpecifier.StartsWith("Date(", StringComparison.InvariantCultureIgnoreCase))
            {
                var format = formatSpecifier.Substring(5, formatSpecifier.Length - 6);
                return new DateFormatter(format);
            }

            throw new NotSupportedException($"Unsupported format specifier: {formatSpecifier}");
        }
    }
}
