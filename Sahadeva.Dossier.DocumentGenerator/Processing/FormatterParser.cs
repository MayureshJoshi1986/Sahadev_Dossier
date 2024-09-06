using Sahadeva.Dossier.DocumentGenerator.Formatting;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class FormatterParser
    {
        private readonly FormatterFactory _formatterFactory;

        public FormatterParser(FormatterFactory formatterFactory)
        {
            _formatterFactory = formatterFactory;
        }

        internal IValueFormatter GetFormatter(string placeholder)
        {
            var formatterPattern = new Regex(@"\|(?<Formatter>.+)\]$", RegexOptions.Compiled);
            var formatMatch = formatterPattern.Match(placeholder);
            if (formatMatch.Success)
            {
                return _formatterFactory.GetFormatter(formatMatch.Groups["Formatter"].Value);
            }

            return new NoOpFormatter();
        }
    }
}
