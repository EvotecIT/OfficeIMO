using System;
using System.Reflection;

namespace PSWriteOffice.Services;

internal static class OpenXmlValueParser
{
    public static bool TryParse<T>(string? value, out T parsed)
    {
        parsed = default!;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var type = typeof(T);
        if (type.IsEnum)
        {
            try
            {
                parsed = (T)Enum.Parse(type, value, ignoreCase: true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        var property = type.GetProperty(value, BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);
        if (property == null)
        {
            return false;
        }

        var resolved = property.GetValue(null);
        if (resolved is T typed)
        {
            parsed = typed;
            return true;
        }

        return false;
    }
}
