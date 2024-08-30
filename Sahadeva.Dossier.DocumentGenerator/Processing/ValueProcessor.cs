using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal class ValueProcessor : IPlaceholderProcessor
    {
        private readonly Text _placeholder;

        public string Expression => _placeholder.Text;

        public ValueProcessor(Text placeholder)
        {
            _placeholder = placeholder;
        }

        public bool Validate()
        {
            throw new NotImplementedException();
        }

        public void ReplacePlaceholder(WordprocessingDocument wordDoc)
        {
            throw new NotImplementedException();
        }
    }
}
