using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.Parser.Importers
{
    public abstract class CsvImporter<TImportValueModel, TImportConfiguration, TMapper> : ICSVImporter<TImportValueModel, TImportConfiguration, TMapper>
        where TImportConfiguration : ImportConfiguration
        where TMapper : ClassMap<TImportValueModel>
    {
        #region Implementation of IImporter<out TImportValueModel,in TImportConfiguration>

        public virtual IEnumerable<TImportValueModel> GetValues(TImportConfiguration configuration, FileDataType fileDataType, Stream dataStream)
        {
            using (StreamReader streamReader = LoadData(fileDataType, dataStream))
            {
                using (var csvReader = new CsvReader(streamReader))
                {
                    csvReader.Configuration.HasHeaderRecord = configuration.FirstRowHasHeaders;
                    csvReader.Configuration.RegisterClassMap<TMapper>();
                    csvReader.Configuration.IgnoreBlankLines = true;
                    csvReader.Configuration.TrimOptions = TrimOptions.Trim | TrimOptions.InsideQuotes;
                    csvReader.Configuration.ShouldSkipRecord = record =>
                    {
                        return record.All(string.IsNullOrEmpty);
                    };


                    IList<TImportValueModel> results = new List<TImportValueModel>();

                    while (csvReader.Read())
                    {
                        TImportValueModel model = csvReader.GetRecord<TImportValueModel>();
                        model = SetDefaults(model);
                        model = ValidateModel(model);
                        results.Add(model);
                    }

                    return results;
                }
            }
        }

        #endregion

        protected virtual StreamReader LoadData(FileDataType fileDataType, Stream dataStream)
        {
            switch (fileDataType)
            {
                case FileDataType.Csv:
                    return new StreamReader(dataStream);
                default:
                    throw new InvalidOperationException(string.Format("The data type {0} cannot be processed by this importer", fileDataType));
            }
        }

        protected abstract TImportValueModel MapDataRowToModel(TImportConfiguration configuration, CsvReader csvReader);

        protected virtual TImportValueModel SetDefaults(TImportValueModel model)
        {
            // By default we do nothing
            return model;
        }

        protected virtual TImportValueModel ValidateModel(TImportValueModel model)
        {
            // By default we do nothing
            return model;
        }

        protected bool HasHeader(string[] headers, string value)
        {
            return headers.Any(x => x.Trim().Equals(value, StringComparison.OrdinalIgnoreCase));
        }
    }
}
