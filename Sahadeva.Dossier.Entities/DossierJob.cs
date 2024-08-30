using System;

namespace Sahadeva.Dossier.Entities
{
    public class DossierJob
    {
        public DateTime CreatedAt { get; set; }

        public int CoverageDossierId { get; set; }

        public string TemplateName { get; set; }
    }
}
