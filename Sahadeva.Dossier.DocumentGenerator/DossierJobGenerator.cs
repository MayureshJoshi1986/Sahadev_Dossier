using Sahadeva.Dossier.DAL;
using Sahadeva.Dossier.Entities;
using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator
{
    // TODO: Move this class into its own project so that we can queue up jobs for processing
    internal class DossierJobGenerator
    {
        public static DossierJob GetJob(string[] args)
        {
            dossierDAL dal = new dossierDAL();
            DataTable lstCDID = dal.FetchPending_DCIDsToProcess_All();

            // TODO: Push jobs to a queue, for now just return the first available job
            return new DossierJob
            {
                CoverageDossierId = Convert.ToInt32(lstCDID.Rows[0]["CDID"]),
                TemplateName = Convert.ToString(lstCDID.Rows[0]["TemplateName"]),
                CreatedAt = DateTime.Now
            };
        }
    }
}
