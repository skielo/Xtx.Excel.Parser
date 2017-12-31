using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.Parser.Importers
{
    public abstract class ExcelImporter<TImportValueModel, TImportConfiguration> : IImporter<TImportValueModel, TImportConfiguration>
        where TImportConfiguration : ImportConfiguration
    {
        public virtual IEnumerable<TImportValueModel> GetValues(TImportConfiguration configuration, FileDataType fileDataType, Stream dataStream)
        {
            using (var excelDataReader = LoadData(fileDataType, dataStream))
            {
                IEnumerable<DataRow> dataRows = GetData(excelDataReader, configuration.WorksheetNames, configuration.FirstRowHasHeaders);

                IEnumerable<TImportValueModel> results = dataRows
                    .Select(dataRow => MapDataRowToModel(configuration, fileDataType, dataRow))
                    .Select(SetDefaults)
                    .Select(ValidateModel)
                    .ToList();

                return results;
            }
        }

        protected virtual IExcelDataReader LoadData(FileDataType fileDataType, Stream dataStream)
        {
            switch (fileDataType)
            {
                case FileDataType.Xls:
                    return ExcelReaderFactory.CreateBinaryReader(dataStream);
                case FileDataType.Xlsx:
                    return ExcelReaderFactory.CreateOpenXmlReader(dataStream);
                default:
                    throw new InvalidOperationException(string.Format("The data type {0} cannot be processed by this importer", fileDataType));
            }
        }

        /// <param name="worksheetFilterSet">A collection of worksheet names to filter the list to.</param>
        protected virtual IEnumerable<string> GetWorksheetNames(IExcelDataReader excelDataReader, IEnumerable<string> worksheetFilterSet = null)
        {
            DataSet workbook = excelDataReader.AsDataSet();

            IEnumerable<string> sheets = workbook.Tables
                .Cast<DataTable>()
                .Select(sheet => sheet.TableName)
                .ToList();

            if (worksheetFilterSet != null)
                sheets = sheets.Where(worksheetFilterSet.Contains);

            return sheets;
        }

        /// <param name="worksheetFilterSet">A collection of worksheet names to filter the list to.</param>
        protected virtual IEnumerable<DataRow> GetData(IExcelDataReader excelDataReader, IEnumerable<string> worksheetFilterSet = null, bool firstRowIsColumnNames = true)
        {

            DataSet workbook = excelDataReader
                .AsDataSet(new ExcelDataSetConfiguration()
                            {
                                UseColumnDataType = true,
                                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = firstRowIsColumnNames
                                }
                            });

            IEnumerable<DataTable> worksheets = workbook
                .Tables
                .Cast<DataTable>()
                .Where(sheet => GetWorksheetNames(excelDataReader, worksheetFilterSet).Contains(sheet.TableName))
                .ToList();

            return worksheets.SelectMany
            (
                worksheet =>
                    worksheet
                        .Rows
                        .Cast<DataRow>()
            );
        }

        protected abstract TImportValueModel MapDataRowToModel(TImportConfiguration configuration, FileDataType fileDataType, DataRow dataRow);

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
    }
}
