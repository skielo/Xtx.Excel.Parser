using System;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace Xtx.Excel.Parser.Importers
{
    public static class EnumHelper
    {
        public static Enum StringToEnum(Type enumType, string value)
        {
            string[] enumValues = GetAllEnumValues(enumType);
            if (enumValues.Contains(value))
                return (Enum)Enum.Parse(enumType, value, true);

            foreach (string enumValue in enumValues)
            {
                MemberInfo[] memberInfo = enumType.GetMember(enumValue);
                if (memberInfo.Any())
                {
                    var enumMemberAttribute = memberInfo[0].GetCustomAttributes(typeof(EnumMemberAttribute), false)
                        .SingleOrDefault()
                        as EnumMemberAttribute;
                    if (enumMemberAttribute != null)
                    {
                        if (enumMemberAttribute.Value == value)
                            return (Enum)Enum.Parse(enumType, enumValue, true);
                    }
                }
            }

            throw new ArgumentException(string.Format("The value '{0}' is not part of the enum '{1}'", value, enumType.Name), "value");
        }

        public static string[] GetAllEnumValues(Type enumType)
        {
            return Enum.GetNames(enumType);
        }
    }
}
