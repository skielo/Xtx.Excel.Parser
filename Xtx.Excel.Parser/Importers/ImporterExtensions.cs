using CsvHelper;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Threading;
using Xtx.Excel.Parser.Configuration;
using Xtx.Excel.Parser.Exceptions;

namespace Xtx.Excel.Parser.Importers
{
    public static class ImporterExtensions
    {
        public static string GetStringValue(this DataRow source, bool firstRowHasHeaders, string columnName, int? columnIndex)
        {
            if (firstRowHasHeaders)
            {
                if (!string.IsNullOrWhiteSpace(columnName))
                    return source[columnName].ToString();
            }
            else
            {
                if (columnIndex != null)
                    return source[columnIndex.Value].ToString();
            }

            return null;
        }

        public static DateTime? GetDateTimeValue(this DataRow source, bool firstRowHasHeaders, string columnName, int? columnIndex, out object rawValue, out string stringValue)
        {
            if (firstRowHasHeaders)
            {
                rawValue = source[columnName];
                if (rawValue is DateTime)
                {
                    stringValue = null;
                    return (DateTime)rawValue;
                }
                stringValue = rawValue.ToString();
            }
            else
            {
                if (columnIndex == null)
                    throw new ArgumentNullException("columnIndex");

                rawValue = source[columnIndex.Value];
                if (rawValue is DateTime)
                {
                    stringValue = null;
                    return (DateTime)rawValue;
                }
                stringValue = rawValue.ToString();
            }

            return null;
        }

        public static string GetStringValue(this CsvReader source, bool firstRowHasHeaders, string columnName, int? columnIndex)
        {
            if (firstRowHasHeaders)
            {
                if (!string.IsNullOrWhiteSpace(columnName))
                    return source.GetField(columnName) ?? string.Empty;
            }
            else
            {
                if (columnIndex != null)
                    return source.GetField(columnIndex.Value) ?? string.Empty;
            }

            return null;
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, string>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
                SetField(propertySetter, target, stringValue);
        }

        /// <remarks>
        /// Ideally we'd type the <see cref="Enum"/> to TEnum but since you cannot restrict generic types to <see cref="Enum"/> we can't.
        /// The result of this is that boxing occurs in the <paramref name="propertySetter"/> when we try to reflect the setter method.
        /// </remarks>
        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, Enum>> propertySetter, Func<string, string> localisedValueCleanerFunction = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            // Null means we're not importing that data, empty string means we can code the value cleaner to deal with them.
            if (stringValue != null)
            {
                if (localisedValueCleanerFunction != null)
                    stringValue = localisedValueCleanerFunction(stringValue);

                MethodInfo setMethod = GetSetMethod(propertySetter);
                Enum enumValue = EnumHelper.StringToEnum(setMethod.GetParameters().Single().ParameterType, stringValue);
                SetField(propertySetter, target, enumValue);
            }
        }

        public static void SetField<TSource, TDestination>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, TDestination>> propertySetter, Func<string, TDestination> localisedValueCleanerFunction = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (localisedValueCleanerFunction != null)
                {
                    TDestination value = localisedValueCleanerFunction(stringValue);

                    SetField(propertySetter, target, value);
                }
            }
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, bool>> propertySetter, Func<string, bool?> customEvaluator = null, bool? defaultValue = null)
        {
            string stringValue = null;
            try
            {
                stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            }
            catch (CsvHelper.MissingFieldException)
            {
                if (defaultValue == null)
                    throw;

                SetField(propertySetter, target, defaultValue.Value);
                return;
            }
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                bool booleanValue;
                if (bool.TryParse(stringValue, out booleanValue))
                    SetField(propertySetter, target, booleanValue);
                else
                {
                    if (customEvaluator != null)
                    {
                        bool? evaluatedBooleanValue = customEvaluator(stringValue) ?? defaultValue;
                        if (evaluatedBooleanValue != null)
                            SetField(propertySetter, target, evaluatedBooleanValue.Value);
                        return;
                    }
                    switch (stringValue.ToLowerInvariant())
                    {
                        case "1":
                        case "true":
                        case "yes":
                            SetField(propertySetter, target, true);
                            break;
                        case "0":
                        case "false":
                        case "no":
                            SetField(propertySetter, target, false);
                            break;
                    }
                }
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, DateTime>> propertySetter, DateTime? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            DateTime dateTimeValue;
            if (string.IsNullOrWhiteSpace(stringValue) && defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
            else if (DateTime.TryParse(stringValue, out dateTimeValue))
                SetField(propertySetter, target, dateTimeValue);
            else
                throw new InvalidDateTimeValueException<TSource, DateTime>(stringValue, propertySetter);
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }
        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, DateTime?>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            DateTime dateTimeValue;
            if (string.IsNullOrWhiteSpace(stringValue))
                SetField(propertySetter, target, null);
            else if (DateTime.TryParse(stringValue, out dateTimeValue))
                SetField(propertySetter, target, dateTimeValue);
            else
                throw new InvalidDateTimeValueException<TSource, DateTime?>(stringValue, propertySetter);
            Thread.CurrentThread.CurrentCulture = originalCulture;
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, decimal>> propertySetter, decimal? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                decimal decimalValue;
                if (decimal.TryParse(stringValue, out decimalValue))
                {
                    if (isPercentage)
                        decimalValue = decimalValue / 100;
                    SetField(propertySetter, target, decimalValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, decimal>(stringValue, propertySetter);
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }
        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, decimal?>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                decimal decimalValue;
                if (string.IsNullOrWhiteSpace(stringValue))
                    SetField(propertySetter, target, null);
                else if (decimal.TryParse(stringValue, out decimalValue))
                {
                    if (isPercentage)
                        decimalValue = decimalValue / 100;
                    SetField(propertySetter, target, decimalValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, decimal?>(stringValue, propertySetter);
            }
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, short>> propertySetter, short? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                short shortValue;
                if (short.TryParse(stringValue, out shortValue))
                {
                    if (isPercentage)
                        shortValue = Convert.ToInt16(shortValue / 100);
                    SetField(propertySetter, target, shortValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, short>(stringValue, propertySetter);
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }

        public static void SetField<TSource>(this CsvReader source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, short?>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                short shortValue;
                if (string.IsNullOrWhiteSpace(stringValue))
                    SetField(propertySetter, target, null);
                else if (short.TryParse(stringValue, out shortValue))
                {
                    if (isPercentage)
                        shortValue = Convert.ToInt16(shortValue / 100);
                    SetField(propertySetter, target, shortValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, short?>(stringValue, propertySetter);
            }
        }

        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, string>> propertySetter, Func<string, string> localisedValueCleanerFunction = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (localisedValueCleanerFunction != null)
                    stringValue = localisedValueCleanerFunction(stringValue);

                SetField(propertySetter, target, stringValue);
            }
        }

        public static void SetField<TSource, TDestination>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, TDestination>> propertySetter, Func<string, TDestination> localisedValueCleanerFunction = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (localisedValueCleanerFunction != null)
                {
                    TDestination value = localisedValueCleanerFunction(stringValue);

                    SetField(propertySetter, target, value);
                }
            }
        }

        public static void SetFieldEnum<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, Enum>> propertySetter, Func<string, string> localisedValueCleanerFunction = null, bool skipEmptyValidation = false)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue) || skipEmptyValidation)
            {
                if (localisedValueCleanerFunction != null)
                    stringValue = localisedValueCleanerFunction(stringValue);

                MethodInfo setMethod = GetSetMethod(propertySetter);
                Enum enumValue = EnumHelper.StringToEnum(setMethod.GetParameters().Single().ParameterType, stringValue);
                SetField(propertySetter, target, enumValue);
            }
        }

        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, bool>> propertySetter, Func<string, bool?> customEvaluator = null, bool? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);
            // Null means we're not importing that data, empty string means we should be, but maybe not on this particular item
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                bool booleanValue;
                if (bool.TryParse(stringValue, out booleanValue))
                    SetField(propertySetter, target, booleanValue);
                else
                {
                    if (customEvaluator != null)
                    {
                        bool? evaluatedBooleanValue = customEvaluator(stringValue) ?? defaultValue;
                        if (evaluatedBooleanValue != null)
                            SetField(propertySetter, target, customEvaluator(stringValue) ?? evaluatedBooleanValue.Value);
                        return;
                    }
                    switch (stringValue.ToLowerInvariant())
                    {
                        case "true":
                        case "1":
                        case "yes":
                            SetField(propertySetter, target, true);
                            break;
                        case "false":
                        case "0":
                        case "no":
                            SetField(propertySetter, target, false);
                            break;
                    }
                }
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }

        public static void SetField<TSource>(this DataRow source, TSource target, FileDataType fileDataType, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, DateTime>> propertySetter, DateTime? defaultValue = null)
        {
            object rawValue;
            string stringValue;
            DateTime? dateValue = GetDateTimeValue(source, firstRowHasHeaders, columnName, columnIndex, out rawValue, out stringValue);
            DateTime dateTimeValue;
            if (dateValue != null)
            {
                dateTimeValue = dateValue.Value;
            }
            else
            {
                CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                switch (fileDataType)
                {
                    case FileDataType.Xls:
                        try
                        {
                            dateTimeValue = DateTime.FromOADate(double.Parse(stringValue));
                        }
                        catch (FormatException)
                        {
                            try
                            {
                                if (!DateTime.TryParse(stringValue, out dateTimeValue))
                                    throw new InvalidDateTimeValueException<TSource, DateTime>(stringValue, propertySetter);
                            }
                            catch
                            {
                                throw new InvalidDateTimeValueException<TSource, DateTime>(stringValue, propertySetter);
                            }
                        }
                        catch
                        {
                            throw new InvalidDateTimeValueException<TSource, DateTime>(stringValue, propertySetter);
                        }
                        break;
                    default:
                        if (!DateTime.TryParse(stringValue, out dateTimeValue))
                            throw new InvalidDateTimeValueException<TSource, DateTime>(stringValue, propertySetter);
                        break;
                }
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }

            SetField(propertySetter, target, dateTimeValue);
        }
        public static void SetField<TSource>(this DataRow source, TSource target, FileDataType fileDataType, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, DateTime?>> propertySetter)
        {
            object rawValue;
            string stringValue;
            DateTime? dateValue = GetDateTimeValue(source, firstRowHasHeaders, columnName, columnIndex, out rawValue, out stringValue);
            DateTime dateTimeValue;
            if (dateValue != null)
            {
                dateTimeValue = dateValue.Value;
            }
            else
            {
                CultureInfo originalCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                switch (fileDataType)
                {
                    case FileDataType.Xls:
                        try
                        {
                            dateTimeValue = DateTime.FromOADate(double.Parse(stringValue));
                        }
                        catch (FormatException)
                        {
                            try
                            {
                                if (!DateTime.TryParse(stringValue, out dateTimeValue))
                                    throw new InvalidDateTimeValueException<TSource, DateTime?>(stringValue, propertySetter);
                            }
                            catch
                            {
                                throw new InvalidDateTimeValueException<TSource, DateTime?>(stringValue, propertySetter);
                            }
                        }
                        catch
                        {
                            throw new InvalidDateTimeValueException<TSource, DateTime?>(stringValue, propertySetter);
                        }
                        break;
                    default:
                        if (!DateTime.TryParse(stringValue, out dateTimeValue))
                            throw new InvalidDateTimeValueException<TSource, DateTime?>(stringValue, propertySetter);
                        break;
                }
                Thread.CurrentThread.CurrentCulture = originalCulture;
            }

            SetField(propertySetter, target, dateTimeValue);
        }

        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, decimal>> propertySetter, decimal? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                decimal decimalValue;
                if (decimal.TryParse(stringValue, out decimalValue))
                {
                    if (isPercentage)
                        decimalValue = decimalValue / 100;
                    SetField(propertySetter, target, decimalValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, decimal>(stringValue, propertySetter);
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }
        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, decimal?>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                decimal decimalValue;
                if (string.IsNullOrWhiteSpace(stringValue))
                    SetField(propertySetter, target, null);
                else if (decimal.TryParse(stringValue, out decimalValue))
                {
                    if (isPercentage)
                        decimalValue = decimalValue / 100;
                    SetField(propertySetter, target, decimalValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, decimal?>(stringValue, propertySetter);
            }
        }

        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, short>> propertySetter, short? defaultValue = null)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                short shortValue;
                if (short.TryParse(stringValue, out shortValue))
                {
                    if (isPercentage)
                        shortValue = Convert.ToInt16(shortValue / 100);
                    SetField(propertySetter, target, shortValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, short>(stringValue, propertySetter);
            }
            else if (defaultValue != null)
                SetField(propertySetter, target, defaultValue.Value);
        }
        public static void SetField<TSource>(this DataRow source, TSource target, bool firstRowHasHeaders, string columnName, int? columnIndex, Expression<Func<TSource, short?>> propertySetter)
        {
            string stringValue = source.GetStringValue(firstRowHasHeaders, columnName, columnIndex);

            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                if (stringValue.StartsWith("$"))
                    stringValue = stringValue.Substring(1);
                bool isPercentage = false;
                if (stringValue.EndsWith("%"))
                {
                    isPercentage = true;
                    stringValue = stringValue.Substring(0, stringValue.Length - 1);
                }

                short shortValue;
                if (string.IsNullOrWhiteSpace(stringValue))
                    SetField(propertySetter, target, null);
                else if (short.TryParse(stringValue, out shortValue))
                {
                    if (isPercentage)
                        shortValue = Convert.ToInt16(shortValue / 100);
                    SetField(propertySetter, target, shortValue);
                }
                else
                    throw new InvalidNumericValueException<TSource, short?>(stringValue, propertySetter);
            }
        }

        private static void SetField<TSource, TProperty>(Expression<Func<TSource, TProperty>> predicate, TSource source, TProperty value)
        {
            MethodInfo setMethod = GetSetMethod(predicate);
            setMethod.Invoke(source, new object[] { value });
        }

        private static MethodInfo GetSetMethod<TSource, TProperty>(Expression<Func<TSource, TProperty>> predicate)
        {
            var member = predicate.Body as MemberExpression;

            // This happens if boxing has occured because the caller didn't know the exact return type.
            // see http://stackoverflow.com/questions/3567857/why-are-some-object-properties-unaryexpression-and-others-memberexpression for more details
            if (member == null)
            {
                var unaryExpression = predicate.Body as UnaryExpression;
                if (unaryExpression != null)
                    member = unaryExpression.Operand as MemberExpression;
            }

            if (member != null)
                return ((PropertyInfo)member.Member).GetSetMethod();

            return null;
        }
    }
}
