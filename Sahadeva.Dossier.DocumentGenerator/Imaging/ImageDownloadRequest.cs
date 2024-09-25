using DocumentFormat.OpenXml.Drawing;

namespace Sahadeva.Dossier.DocumentGenerator.Imaging
{
    internal class ImageDownloadRequest
    {
        public string ImageUrl { get; private set; }

        public Blip Blip { get; private set; }

        public ImageDownloadRequest(string imagePlaceholder, Blip blip)
        {
            ImageUrl = ParseImagePlaceholder(imagePlaceholder);
            Blip = blip;
        }

        private static string ParseImagePlaceholder(string imagePlaceholder)
        {
            return imagePlaceholder.Replace("AF.Image=", string.Empty);
        }
    }
}
