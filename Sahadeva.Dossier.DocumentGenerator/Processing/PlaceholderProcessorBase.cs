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

        public string DataSourceName { get; protected set; } = string.Empty;

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            PlaceholderOptionsRegex = GetPlaceholderOptionsRegex();
            ExtractPlaceholderOptions();

            if (string.IsNullOrWhiteSpace(DataSourceName)) { throw new ApplicationException("DataSourceName must be set while creating a processor before exiting 'ExtractPlaceholderOptions'."); }
        }

        public abstract void ReplacePlaceholder(WordprocessingDocument wordDoc, DataTable data);

        protected abstract Regex GetPlaceholderOptionsRegex();

        protected abstract void ExtractPlaceholderOptions();
    }
}
