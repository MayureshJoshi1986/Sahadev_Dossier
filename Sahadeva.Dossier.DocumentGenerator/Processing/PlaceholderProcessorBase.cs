using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal abstract class PlaceholderProcessorBase : IPlaceholderProcessor
    {
        protected readonly Regex PlaceholderRegex;

        protected Text Placeholder { get; private set; }

        public string Expression => Placeholder.Text;

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            PlaceholderRegex = CreatePlaceholderRegex();
            ExtractPlaceholderParams();
        }

        public abstract void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data);

        protected abstract Regex CreatePlaceholderRegex();

        protected abstract void ExtractPlaceholderParams();
    }
}
