using CsvHelper.Configuration;
using Xtx.Excel.ParserTests.Models;

namespace Xtx.Integrations.Excel.Importers.Mappers
{
    /// <summary>
    /// This is the Mapper class to configure the Importer.
    /// </summary>
    public class FakeMapper : ClassMap<FakeImportModel>
    {
        public FakeMapper()
        {
            AutoMap();
        }
    }
}
