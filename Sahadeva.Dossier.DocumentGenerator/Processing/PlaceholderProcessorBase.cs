using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal abstract class PlaceholderProcessorBase<T> : IPlaceholderProcessor<T>
    {
        protected Text Placeholder { get; private set; }

        public PlaceholderProcessorBase(Text placeholder)
        {
            Placeholder = placeholder;
            ParsePlaceholder();
        }

        public abstract void ReplacePlaceholder(T data);

        public abstract void ParsePlaceholder();
    }
}
