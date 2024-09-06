namespace Sahadeva.Dossier.DocumentGenerator.Formatting
{
    internal class NoOpFormatter : IValueFormatter
    {
        public string Format(string value)
        {
            return value;
        }
    }
}
