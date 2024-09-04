using DocumentFormat.OpenXml.Packaging;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    public interface IPlaceholderProcessor
    {
        string DataSourceName { get; }

        string Expression { get; }

        void ReplacePlaceholder(WordprocessingDocument wordDoc, DataTable data);
    }
}