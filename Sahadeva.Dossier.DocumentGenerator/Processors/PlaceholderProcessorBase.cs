using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal abstract class PlaceholderProcessorBase
    {
        protected Text Placeholder { get; private set; }

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            SetPlaceholderOptions();
        }

        /// <summary>
        /// Parse the placeholder and set its options eg. TableName, ColumnName
        /// </summary>
        public abstract void SetPlaceholderOptions();
    }
}
