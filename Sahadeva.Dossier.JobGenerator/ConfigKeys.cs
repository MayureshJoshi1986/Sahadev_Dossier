namespace Sahadeva.Dossier.JobGenerator
{
    internal static class ConfigKeys
    {
        public const string DEBUG_ENV = "DEBUG_ENV";
        public const string ConnectionString = "ConnectionString";
        public const string MinimumLogLevel = "MinimumLogLevel";
        public const string LogPath = "LogPath";
        public const string PollingIntervalInSeconds = "PollingIntervalInSeconds";
        public const string MaxBatchSize = "MaxBatchSize";
        public const string SQSAccessKey = "SQS:AccessKey";
        public const string SQSSecret = "SQS:Secret";
        public const string SQSEndpoint = "SQS:Endpoint";
    }
}
