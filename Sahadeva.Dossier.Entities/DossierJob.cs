using System;

namespace Sahadeva.Dossier.Entities
{
    public class DossierJob
    {
        public DossierJob(string runId, string templateName, int coverageDossierId, string outputFilePath)
        {
            Timestamp = DateTime.UtcNow;
            RunId = runId;
            TemplateName = templateName;
            CoverageDossierId = coverageDossierId;
            OutputFilePath = outputFilePath;
        }

        public string RunId { get; private set; }

        public DateTime Timestamp { get; private set; }

        public int CoverageDossierId { get; private set; }

        public string TemplateName { get; private set; }

        public string OutputFilePath { get; private set; }
    }
}
