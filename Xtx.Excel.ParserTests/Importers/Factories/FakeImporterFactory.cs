using System;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.ParserTests.Importers.Factories
{
    /// <summary>
    /// It's important have a factory to determine the right Importer based on the file extension.
    /// </summary>
    public class FakeImporterFactory
    {
        public static IFakeImporter GetImporter(FileDataType fileDataType)
        {
            switch (fileDataType)
            {
                case FileDataType.Xls:
                case FileDataType.Xlsx:
                    return new FakeExcelImporter();
                case FileDataType.Csv:
                    return new FakeCsvImporter();
                default:
                    throw new InvalidOperationException(string.Format("The data type {0} cannot be processed by this importer", fileDataType));
            }
        }
    }
}
