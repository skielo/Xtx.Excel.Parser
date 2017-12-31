using System;
using System.Linq.Expressions;

namespace Xtx.Excel.Parser.Exceptions
{
    public class InvalidDateTimeValueException<TSource, TProperty> : ExcelImportException
    {
        public InvalidDateTimeValueException(string value, Expression<Func<TSource, TProperty>> property)
            : base(string.Format("The value of '{0}' for '{1}' could not be converted into a date.", value, property.GetPropertyDescription()))
        {
        }
    }
}
