using System.Data;

namespace Sahadeva.Dossier.DocumentGenerator.Processing
{
    internal interface IPlaceholderWithDataSource : IPlaceholderProcessor<DataTable>
    {
        string TableName { get;  }
    }
}
