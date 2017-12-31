using Xtx.Excel.Parser.Configuration;

namespace Xtx.Excel.ParserTests.Configuration
{
    public class FakeImportConfiguration : ImportConfiguration
    {
        public string UserEmailAddressColumnName { get; set; }

        public string FirstNameColumnName { get; set; }

        public string LastNameColumnName { get; set; }

        public int? UserEmailAddressColumnIndex { get; set; }

        public int? FirstNameColumnIndex { get; set; }

        public int? LastNameColumnIndex { get; set; }
        
        public FakeImportConfiguration()
			: this (true)
		{
        }

        public FakeImportConfiguration(bool firstRowHasHeaders)
			: base(firstRowHasHeaders)
		{
            if (FirstRowHasHeaders)
            {
                FirstNameColumnName = "First Name";
                LastNameColumnName = "Last Name";
                UserEmailAddressColumnName = "Email";
            }
            else
            {
                FirstNameColumnIndex = 0;
                LastNameColumnIndex = 1;
                UserEmailAddressColumnIndex = 2;
            }
        }
    }
}
