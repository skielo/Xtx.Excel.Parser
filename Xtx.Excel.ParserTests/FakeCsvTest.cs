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
    public class FakeCsvTest
    {
        [Test]
        public void Can_Read_And_Parse()
        {
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var data = new FileStream($"{dir}\\Data\\FakeCsv.csv", FileMode.Open);
            var importer = FakeImporterFactory.GetImporter(FileDataType.Csv);
            var configuration = new FakeImportConfiguration();

            IEnumerable<FakeImportModel> results = importer.GetValues(configuration, FileDataType.Csv, data);

            Assert.AreEqual(2, results.Count());
            Assert.IsTrue(results.All(x => !string.IsNullOrEmpty(x.FirstName)));
        }
    }
}
