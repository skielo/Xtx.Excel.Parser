using System;
using System.Linq.Expressions;

namespace Xtx.Excel.Parser.Exceptions
{
    public class MissingRequiredValueException<TModel> : ExcelImportException
    {
        public MissingRequiredValueException(Expression<Func<TModel, object>> missingProperty, Expression<Func<TModel, object>> wantingProperty)
            : base(string.Format("The value for '{0}' was not provided but is required by '{1}'.", missingProperty.GetPropertyDescription(), wantingProperty.GetPropertyDescription()))
        {
        }

        public MissingRequiredValueException(params Expression<Func<TModel, object>>[] missingProperties)
            : base(string.Format("The value for {0} was not provided.", missingProperties.GetMissingPropertyDescriptions()))
        {
        }

        public MissingRequiredValueException(string message)
            : base(message)
        {
        }
    }

    public static class MissingRequiredValueExceptionExtensions
    {
        public static string GetMissingPropertyDescriptions<TModel>(this Expression<Func<TModel, object>>[] missingProperties)
        {
            if (missingProperties == null)
                return "Something";

            string result = string.Empty;
            foreach (Expression<Func<TModel, object>> missingProperty in missingProperties)
            {
                if (!string.IsNullOrWhiteSpace(result))
                    result = string.Concat(result, ", ");
                result = string.Concat(result, "'", missingProperty.GetPropertyDescription(), "'");
            }

            return result;
        }
    }
}
