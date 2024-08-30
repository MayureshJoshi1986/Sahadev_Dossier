using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal abstract class PlaceholderProcessorBase : IPlaceholderProcessor
    {
        protected Text Placeholder { get; private set; }

        public string Expression => Placeholder.Text;

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            ExtractPlaceholderParams();
        }

        public abstract void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data);

        protected abstract void ExtractPlaceholderParams();
    }
}
