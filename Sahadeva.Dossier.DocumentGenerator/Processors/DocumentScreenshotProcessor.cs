using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DocumentGenerator.Imaging;
using Sahadeva.Dossier.DocumentGenerator.OpenXml;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Processors
{
    internal partial class DocumentScreenshotProcessor : PlaceholderProcessorBase<Drawing>, IPlaceholderWithDataSource
    {
        private readonly ScreenshotService _screenshotService;

        public string TableName { get; private set; } = string.Empty;

        protected string ColumnName { get; private set; } = string.Empty;

        public DocumentScreenshotProcessor(Placeholder<Drawing> placeholder, ScreenshotService screenshotService) : base(placeholder)
        {
            _screenshotService = screenshotService;
        }

        public void ReplacePlaceholder(DataTable data)
        {
            var value = GetValueFromSource(data);
            var screenshotUrl = _screenshotService.GetScreenshotUrl(value);
            SetImageUrl(screenshotUrl);
        }

        private void SetImageUrl(string url)
        {
            var nonVisualProps = Placeholder.Element.Descendants<DocProperties>().First();
            nonVisualProps.Description = $"AF.Image={url}";
        }

        protected string GetValueFromSource(DataTable data)
        {
            if (data.Rows.Count != 1) { throw new ApplicationException($"Attempt to use a single value placeholder '{Placeholder.Text}' for multiple possible values"); }

            if (!data.Columns.Contains(ColumnName)) { throw new ApplicationException($"Could not find column '{ColumnName}' in '{TableName}"); }

            return data.Rows[0][ColumnName].ToString()!;
        }

        public override void SetPlaceholderOptions()
        {
            var match = OptionsRegex().Match(Placeholder.Text);
            if (match.Success)
            {
                TableName = match.Groups["TableName"].Value;
                ColumnName = match.Groups["ColumnName"].Value;
            }
            else
            {
                throw new ApplicationException($"Could not parse {Placeholder.Text}");
            }
        }

        [GeneratedRegex(@"\[AF\.Screenshot:(?<TableName>[^\.\]]+)\.(?<ColumnName>[^\|\]]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled)]
        private static partial Regex OptionsRegex();
    }
}
