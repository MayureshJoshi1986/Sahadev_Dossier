using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class PlaceholderFactory
    {
        private static readonly Dictionary<string, Func<Text, IPlaceholderProcessor>> _processorFactory =
            new()
            {
                { "[AF.Value:", placeholder => new ValueProcessor(placeholder) },
                { "[AF.MultilineValue:", placeholder => new MultilineValueProcessor(placeholder) },
                { "[AF.Table:", placeholder => new TableProcessor(placeholder) }
            };

        public IPlaceholderProcessor CreateProcessor(Text placeholder)
        {
            foreach (var entry in _processorFactory)
            {
                if (placeholder.Text.StartsWith(entry.Key))
                {
                    return entry.Value(placeholder);
                }
            }

            throw new NotSupportedException($"Unsupported placeholder type: {placeholder.Text}");
        }
    }
}
