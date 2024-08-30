using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class PlaceholderFactory
    {
        public IPlaceholderProcessor CreateProcessor(Text placeholder)
        {
            if (placeholder.Text.StartsWith("[AF.Value:"))
            {
                return new ValueProcessor(placeholder);
            }
            else
            {
                throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}");
            }
        }
    }
}
