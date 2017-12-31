using System;

namespace Xtx.Excel.Parser.Exceptions
{
    public class InvalidImportDate : Exception
    {
        public InvalidImportDate(string message)
            : base(message)
        {

        }
    }
}
