using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.Common.Configuration;

namespace Sahadeva.Dossier.DocumentGenerator.Imaging
{
    internal class ImageDownloader
    {
        private readonly int _imageMaxDegreeOfParallelism;
        private static readonly HttpClient _httpClient = new();
        private readonly object _documentLock = new();
        private const int DEFAULT_MAX_DEGREE_OF_PARALLELISM = 10;

        public ImageDownloader()
        {
            _imageMaxDegreeOfParallelism = int.Parse(ConfigurationManager.Settings["ImageMaxDegreeOfParallelism"] ?? DEFAULT_MAX_DEGREE_OF_PARALLELISM.ToString());
        }

        /// <summary>
        /// Downloads images in parallel to speed up document processing.
        /// </summary>
        internal async Task DownloadImagesAsync(WordprocessingDocument document)
        {
            var drawings = document.MainDocumentPart!.Document.Descendants<Drawing>().ToList();
            var imageRequests = new List<ImageDownloadRequest>();
            foreach (var drawing in drawings)
            {
                var nonVisualProps = drawing.Descendants<DocProperties>().FirstOrDefault();

                if (nonVisualProps != null && nonVisualProps.Description != null &&
                    nonVisualProps.Description.Value!.StartsWith("AF.Image="))
                {
                    var blip = drawing.Descendants<Blip>().FirstOrDefault();
                    if (blip != null)
                    {
                        imageRequests.Add(new ImageDownloadRequest(nonVisualProps.Description.Value, blip));
                    }
                }
            }

            await ReplaceImagesAsync(document, imageRequests);
        }

        private async Task ReplaceImagesAsync(WordprocessingDocument document, IEnumerable<ImageDownloadRequest> imageRequests)
        {
            using (var semaphore = new SemaphoreSlim(_imageMaxDegreeOfParallelism))
            {
                var downloadTasks = imageRequests.Select(async request =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        byte[] imageData = await _httpClient.GetByteArrayAsync(request.ImageUrl);
                        // OpenXml document modifications are not thread safe. Ensure only one thread is modifying the document at any given point
                        // the real bottleneck would be the image downloads which we are running in parallel
                        lock (_documentLock)
                        {
                            ReplaceImageInDocument(document, request.Blip, imageData);
                        }
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                // Wait for all downloads to complete
                await Task.WhenAll(downloadTasks);
            }
        }

        private void ReplaceImageInDocument(WordprocessingDocument document, Blip blip, byte[] imageData)
        {
            // Retrieve the existing image part
            var oldImagePart = document.MainDocumentPart!.GetPartById(blip.Embed!.Value!) as ImagePart;

            // Add the new image part (Ensure the correct image type is used here)
            ImagePart newImagePart = document.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (var imageStream = new MemoryStream(imageData))
            {
                newImagePart.FeedData(imageStream);
            }

            // Update the relationship ID to the new image part
            blip.Embed = document.MainDocumentPart.GetIdOfPart(newImagePart);

            // Delete the old image part if it exists
            if (oldImagePart != null)
            {
                document.MainDocumentPart.DeletePart(oldImagePart);
            }
        }
    }
}
