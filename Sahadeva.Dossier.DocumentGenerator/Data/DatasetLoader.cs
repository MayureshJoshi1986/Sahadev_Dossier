using DocumentFormat.OpenXml.Wordprocessing;
using Sahadeva.Dossier.DAL;
using Sahadeva.Dossier.DocumentGenerator.Extensions;
using Sahadeva.Dossier.Entities;
using System.Data;
using System.Text.RegularExpressions;

namespace Sahadeva.Dossier.DocumentGenerator.Data
{
    internal partial class DatasetLoader
    {
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

            var requiredDataSources = GetUniqueDataSources(placeholders);

            foreach (var dataSource in Enum.GetValues<DossierDataSet>())
            {
                if (requiredDataSources.Contains(dataSource.ToString()))
                {
                    var data = _dal.FetchData(job.CoverageDossierId, dataSource);
                    dataset.AddTableToDataSet(data, dataSource.ToString());
                }
            }

            return dataset;
        }

        private HashSet<string> GetUniqueDataSources(List<Text> placeholders)
        {
            var tableNames = new HashSet<string>();

            foreach (var placeholder in placeholders)
            {
                var match = DataSourceRegex().Match(placeholder.Text);

                if (match.Success)
                {
                    tableNames.Add(match.Groups["TableName"].Value);
                }
            }

            return tableNames;
        }

        [GeneratedRegex(@"\[AF\.[^\:]+\:(?<TableName>[^\.\;\]]+)", RegexOptions.Compiled)]
        private static partial Regex DataSourceRegex();
    }
}
