using System;


namespace Xtx.Excel.Parser.Exceptions
{
    public class ExcelImportException : Exception
    {
        public ExcelImportException(string message)
            : base(message)
        { }
    }
}
