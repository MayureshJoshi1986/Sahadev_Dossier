using Microsoft.Extensions.Options;
using Sahadeva.Dossier.DocumentGenerator.Configuration;

namespace Sahadeva.Dossier.DocumentGenerator.Imaging
{
    internal class GraphService
    {
        private readonly GraphOptions _options;

        public GraphService(IOptions<GraphOptions> options)
        {
            _options = options.Value;
        }

        internal string GetGraphUrl(int cdid, string graphType)
        {
            // TODO: Would have liked to use the same approach as the screenshot url service but for some reason 
            // the graph endpoint does not seem to like the way the params are encoded.
            // Cannot debug without help from AdFactors so leaving it for now
            return _options.Endpoint
                .Replace("#CDIDVALUE#", cdid.ToString())
                .Replace("#METADATAVALUE#", graphType);
        }
    }
}
