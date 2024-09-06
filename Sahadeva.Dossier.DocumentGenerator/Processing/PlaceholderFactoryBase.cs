using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    /// <summary>
    /// Base class which can be used to create specialised factories for elements that are specific to a given context
    /// eg. placeholders within a table, section etc
    /// </summary>
    internal abstract class PlaceholderFactoryBase<T>
    {
        internal abstract T CreateProcessor(Text placeholder);

        /// <summary>
        /// Extract the placeholder type from the placeholder expressions.
        /// </summary>
        /// <param name="placeholder"></param>
        /// <returns></returns>
        protected string GetPlaceholderType(string placeholder)
        {
            var placeholderTypePattern = new Regex(@"\[AF\.(?<Type>[^\[\]:]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var match = placeholderTypePattern.Match(placeholder);
            if (match.Success)
            {
                return match.Groups["Type"].Value;
            }

            return string.Empty;
        }
    }
}
