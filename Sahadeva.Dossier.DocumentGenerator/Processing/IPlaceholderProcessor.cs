using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    public interface IPlaceholderProcessor
    {
        string Expression { get; }

        void ReplacePlaceholder(WordprocessingDocument wordDoc, DataSet data);
    }
}