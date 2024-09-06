namespace Sahadeva.Dossier.DocumentGenerator.Formatting
{
    internal class UrlFormatter : IValueFormatter
    {
        private readonly string _url;

        public UrlFormatter(string url)
        {
            _url = url;
        }

        public string Format(string value)
        {
            // TODO: Check what a link looks like in the word doc
            return $"<a href=\"{_url}\">{value}</a>";
        }
    }
}
