using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xtx.Excel.ParserTests.Configuration;
using Xtx.Excel.ParserTests.Importers.Factories;
using Xtx.Excel.ParserTests.Models;
using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.ParserTests
{
    [TestFixture]
    public class FakeExcelTest
    {
        [Test]
        public void Can_Read_And_Parse()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var data = new FileStream($"{dir}\\Data\\FakeExcel.xls", FileMode.Open);
            var importer = FakeImporterFactory.GetImporter(FileDataType.Xls);
            var configuration = new FakeImportConfiguration();

            IEnumerable<FakeImportModel> results = importer.GetValues(configuration, FileDataType.Xls, data);

            Assert.AreEqual(2, results.Count());
            Assert.IsTrue(results.All(x => !string.IsNullOrEmpty(x.FirstName)));
        }
    }
}
