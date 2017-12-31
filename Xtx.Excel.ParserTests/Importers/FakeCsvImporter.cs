using CsvHelper;
using Xtx.Excel.ParserTests.Configuration;
using Xtx.Excel.ParserTests.Models;
using Xtx.Excel.Parser.Importers;
using Xtx.Integrations.Excel.Importers.Mappers;

namespace Xtx.Excel.ParserTests.Importers
{
    public class FakeCsvImporter : CsvImporter<FakeImportModel, FakeImportConfiguration, FakeMapper>, IFakeImporter
    {
        protected override FakeImportModel MapDataRowToModel(FakeImportConfiguration configuration, CsvReader csvReader)
        {
            var result = new FakeImportModel();

            csvReader.SetField(result, configuration.FirstRowHasHeaders, configuration.FirstNameColumnName, configuration.FirstNameColumnIndex, value => result.FirstName);
            csvReader.SetField(result, configuration.FirstRowHasHeaders, configuration.LastNameColumnName, configuration.LastNameColumnIndex, value => result.LastName);
            csvReader.SetField(result, configuration.FirstRowHasHeaders, configuration.UserEmailAddressColumnName, configuration.UserEmailAddressColumnIndex, value => result.UserEmailAddress);

            return result;
        }
    }
}
