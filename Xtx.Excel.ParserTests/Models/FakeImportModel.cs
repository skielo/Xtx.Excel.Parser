using System.ComponentModel;

namespace Xtx.Excel.ParserTests.Models
{
    /// <summary>
    /// In order to use the library it's important to have a model that matches with the data model
    /// we want to import from the Csv or Excel file.
    /// </summary>
    public class FakeImportModel
    {
        [Description("First Name")]
        public string FirstName { get; set; }

        [Description("Last Name")]
        public string LastName { get; set; }

        [Description("Emaill Address")]
        public string UserEmailAddress { get; set; }
    }
}
