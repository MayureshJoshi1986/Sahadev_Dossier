using Microsoft.Extensions.Options;
using Sahadeva.Dossier.DocumentGenerator.Configuration;
using System.Web;

namespace Sahadeva.Dossier.DocumentGenerator.Imaging
{
    internal class ScreenshotService
    {
        private readonly ScreenshotOptions _options;

        public ScreenshotService(IOptions<ScreenshotOptions> options)
        {
            _options = options.Value;
        }

        internal string GetScreenshotUrl(string articleUrl)
        {
            var screenshotApi = new UriBuilder(_options.Endpoint);

            var query = HttpUtility.ParseQueryString(string.Empty); 

            query["url"] = articleUrl;
            query["width"] = _options.Width.ToString();
            query["height"] = _options.Height.ToString();
            query["delay"] = _options.Delay.ToString();

            screenshotApi.Query = query.ToString();

            return screenshotApi.ToString();
        }
    }
}
