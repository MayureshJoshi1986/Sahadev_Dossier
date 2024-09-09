using DocumentFormat.OpenXml.Wordprocessing;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal abstract class DocumentPlaceholderProcessorBase : PlaceholderProcessorBase, IDocumentPlaceholderProcessor
    {
        public string TableName { get; protected set; } = string.Empty;

        public DocumentPlaceholderProcessorBase(Text placeholder) : base(placeholder)
        {
        }

        /// <summary>
        /// Replace the placeholder with the actual content
        /// </summary>
        /// <param name="data"></param>
        public abstract void ReplacePlaceholder(DataTable data);
    }
}
