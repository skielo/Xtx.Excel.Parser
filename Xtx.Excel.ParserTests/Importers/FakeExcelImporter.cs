using System.Data;
using Xtx.Excel.ParserTests.Configuration;
using Xtx.Excel.ParserTests.Models;
using Xtx.Excel.Parser.Configuration;
using Xtx.Excel.Parser.Importers;

namespace Xtx.Excel.ParserTests.Importers
{
    public class FakeExcelImporter : ExcelImporter<FakeImportModel, FakeImportConfiguration>, IFakeImporter
    {
        protected override FakeImportModel MapDataRowToModel(FakeImportConfiguration configuration, FileDataType fileDataType, DataRow dataRow)
        {
            var result = new FakeImportModel();

            dataRow.SetField(result, configuration.FirstRowHasHeaders, configuration.FirstNameColumnName, configuration.FirstNameColumnIndex, value => result.FirstName);
            dataRow.SetField(result, configuration.FirstRowHasHeaders, configuration.LastNameColumnName, configuration.LastNameColumnIndex, value => result.LastName);
            dataRow.SetField(result, configuration.FirstRowHasHeaders, configuration.UserEmailAddressColumnName, configuration.UserEmailAddressColumnIndex, value => result.UserEmailAddress);

            return result;
        }
    }
}
