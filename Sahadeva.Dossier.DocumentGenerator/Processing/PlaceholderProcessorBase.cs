using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal abstract class PlaceholderProcessorBase : IPlaceholderProcessor
    {
        /// <summary>
        /// Regular expression that will be used to extract the options from the placeholder
        /// </summary>
        protected readonly Regex PlaceholderOptionsRegex;

        protected Text Placeholder { get; private set; }

        public string Expression => Placeholder.Text;

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            PlaceholderOptionsRegex = GetPlaceholderOptionsRegex();
            ExtractPlaceholderOptions();
        }

        public abstract void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data);

        protected abstract Regex GetPlaceholderOptionsRegex();

        protected abstract void ExtractPlaceholderOptions();
    }
}
