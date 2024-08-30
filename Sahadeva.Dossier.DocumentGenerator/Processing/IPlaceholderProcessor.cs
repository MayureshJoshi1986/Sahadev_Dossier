using DocumentFormat.OpenXml.Packaging;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    public interface IPlaceholderProcessor
    {
        string Expression { get; }

        void ReplacePlaceholder(WordprocessingDocument wordDoc);
        bool Validate();
    }
}