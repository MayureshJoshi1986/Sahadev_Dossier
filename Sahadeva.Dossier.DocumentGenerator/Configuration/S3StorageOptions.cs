namespace Sahadeva.Dossier.DocumentGenerator.Configuration
{
    internal class S3StorageOptions
    {
        public string BucketName { get; set; } = string.Empty;
        
        public string TemplatePath { get; set; } = string.Empty;

        public string OutputPath { get; set; } = string.Empty;
    }
}
