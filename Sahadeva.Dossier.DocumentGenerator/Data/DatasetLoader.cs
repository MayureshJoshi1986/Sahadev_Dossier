using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DAL;
using Sahadeva.Dossier.DocumentGenerator.Extensions;
using Sahadeva.Dossier.Entities;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Data
{
    internal class DatasetLoader
    {
        /// <summary>
        /// List of placeholders that will contain datasource information.
        /// Only placeholders listed here will be scanned
        /// </summary>
        private readonly List<Regex> _dataSources =
        [
            new(@"\[AF\.Value:(?<TableName>[^\.]+)\.\w+\]", RegexOptions.Compiled),
            new(@"\[AF\.MultilineValue:(?<TableName>[^\.]+)\.\w+\]", RegexOptions.Compiled),
            new(@"\[AF\.Table:(?<TableName>[^\]]+)\]", RegexOptions.Compiled)
        ];

        private readonly DossierDAL _dal;

        public DatasetLoader(DossierDAL dal)
        {
            _dal = dal;
        }

        /// <summary>
        /// Examines the placeholders and loads the necessary data
        /// </summary>
        /// <param name="job"></param>
        /// <param name="placeholders"></param>
        /// <returns>A DataSet which contains all the data required for building the dossier</returns>
        internal DataSet LoadDataset(DossierJob job, List<Text> placeholders)
        {
            var dataset = new DataSet();
            DataTable data;
            
            var requiredDataSources = GetUniqueDataSources(placeholders);

            if (requiredDataSources.Contains(DossierDataSet.ClientData.ToString()))
            {
                data = _dal.FetchData(job.CoverageDossierId, DossierDataSet.ClientData);
                dataset.AddTableToDataSet(data, DossierDataSet.ClientData.ToString());
            }

            if (requiredDataSources.Contains(DossierDataSet.OverviewTable.ToString()))
            {
                data = _dal.FetchData(job.CoverageDossierId, DossierDataSet.OverviewTable);
                dataset.AddTableToDataSet(data, DossierDataSet.OverviewTable.ToString());
            }

            return dataset;
        }

        private HashSet<string> GetUniqueDataSources(List<Text> placeholders)
        {
            var tableNames = new HashSet<string>();

            foreach (var placeholder in placeholders)
            {
                foreach (var regex in _dataSources)
                {
                    var match = regex.Match(placeholder.Text);

                    if (match.Success)
                    {
                        tableNames.Add(match.Groups["TableName"].Value);
                        break;
                    }
                }
            }

            return tableNames;
        }
    }
}
