namespace Sahadeva.Dossier.DocumentGenerator.Formatting
{
    internal class FormatterFactory
    {
        internal IValueFormatter GetFormatter(string formatSpecifier)
        {
            if (formatSpecifier.StartsWith("Date(", StringComparison.InvariantCultureIgnoreCase))
            {
                var format = formatSpecifier.Substring(5, formatSpecifier.Length - 6);
                return new DateFormatter(format);
            }
            else if (formatSpecifier.StartsWith("URL(", StringComparison.InvariantCultureIgnoreCase))
            {
                var url = formatSpecifier.Substring(4, formatSpecifier.Length - 5);
                return new UrlFormatter(url);
            }

            throw new NotSupportedException($"Unsupported format specifier: {formatSpecifier}");
        }
    }
}
