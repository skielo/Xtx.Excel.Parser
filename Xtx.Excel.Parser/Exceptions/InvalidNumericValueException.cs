using System;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Reflection;

namespace Xtx.Excel.Parser.Exceptions
{
    public class InvalidNumericValueException<TSource, TProperty> : ExcelImportException
    {
        public InvalidNumericValueException(string value, Expression<Func<TSource, TProperty>> property)
            : base(string.Format("The value of '{0}' for '{1}' could not be converted into a number.", value, property.GetPropertyDescription()))
        {
        }
    }

    public static class InvalidNumericValueExceptionExtensions
    {
        public static string GetPropertyDescription<TSource, TProperty>(this Expression<Func<TSource, TProperty>> property)
        {
            MemberExpression memberExpression;

            var unaryExpression = property.Body as UnaryExpression;
            if (unaryExpression != null)
            {
                memberExpression = unaryExpression.Operand as MemberExpression;
            }
            else
            {
                memberExpression = property.Body as MemberExpression;
            }

            if (memberExpression != null)
            {
                var descriptionAttribute = memberExpression.Member.GetCustomAttribute<DescriptionAttribute>();
                if (descriptionAttribute != null)
                {
                    return descriptionAttribute.Description;
                }
            }

            return null;
        }
    }
}
