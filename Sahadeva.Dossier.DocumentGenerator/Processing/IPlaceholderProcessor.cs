namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    public interface IPlaceholderProcessor<T>
    {
        /// <summary>
        /// Replace the placeholder text with actual content
        /// </summary>
        /// <param name="data"></param>
        void ReplacePlaceholder(T data);
    }
}