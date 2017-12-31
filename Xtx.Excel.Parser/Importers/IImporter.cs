using CsvHelper.Configuration;
using System.Collections.Generic;
using System.IO;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.Parser.Importers
{
    public interface IImporter<out TImportValueModel, in TImportConfiguration>
        where TImportConfiguration : ImportConfiguration
    {
        IEnumerable<TImportValueModel> GetValues(TImportConfiguration configuration, FileDataType fileDataType, Stream dataStream);
    }

    public interface ICSVImporter<out TImportValueModel, in TImportConfiguration, in TMapper>: IImporter<TImportValueModel, TImportConfiguration>
        where TImportConfiguration : ImportConfiguration
        where TMapper : ClassMap<TImportValueModel>
    {
    }
}
