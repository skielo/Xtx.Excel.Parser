using System.Collections.Generic;
using System.IO;
using Xtx.Excel.ParserTests.Configuration;
using Xtx.Excel.ParserTests.Models;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.ParserTests.Importers
{
    /// <summary>
    /// This interface is the definition of the importer in order to get the values.
    /// </summary>
    public interface IFakeImporter
    {
        IEnumerable<FakeImportModel> GetValues(FakeImportConfiguration configuration, FileDataType fileDataType, Stream dataStream);
    }
}
