#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV;

internal static class CsvValueConverter
{
    public static T? ConvertTo<T>(object? value, CultureInfo culture)
    {
        var targetType = typeof(T);
        if (!TryConvert(value, targetType, culture, out var result, out var error))
        {
            throw new CsvException(error ?? $"Value '{value}' cannot be converted to {targetType.Name}");
        }

        return (T?)result;
    }

    public static bool TryConvert(object? value, Type targetType, CultureInfo culture, out object? result, out string? error)
    {
        error = null;
        result = null;

        var underlyingType = Nullable.GetUnderlyingType(targetType);
        var effectiveType = underlyingType ?? targetType;

        if (value is null)
        {
            if (underlyingType is null && targetType.IsValueType)
            {
                error = $"Cannot assign null to non-nullable type {targetType.Name}.";
                return false;
            }

            result = null;
            return true;
        }

        if (effectiveType.IsInstanceOfType(value))
        {
            result = value;
            return true;
        }

        if (value is string s)
        {
            return TryConvertFromString(s, effectiveType, culture, out result, out error);
        }

        try
        {
            result = System.Convert.ChangeType(value, effectiveType, culture);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            result = null;
            return false;
        }
    }

    private static bool TryConvertFromString(string text, Type targetType, CultureInfo culture, out object? result, out string? error)
    {
        error = null;
        result = null;

        try
        {
            if (targetType == typeof(string))
            {
                result = text;
                return true;
            }

            if (targetType == typeof(int))
            {
                result = int.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(long))
            {
                result = long.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(short))
            {
                result = short.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(byte))
            {
                result = byte.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(bool))
            {
                if (bool.TryParse(text, out var b))
                {
                    result = b;
                    return true;
                }

                if (text == "0")
                {
                    result = false;
                    return true;
                }

                if (text == "1")
                {
                    result = true;
                    return true;
                }

                error = $"Cannot parse '{text}' as boolean.";
                return false;
            }

            if (targetType == typeof(double))
            {
                result = double.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(decimal))
            {
                result = decimal.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(float))
            {
                result = float.Parse(text, NumberStyles.Any, culture);
                return true;
            }

            if (targetType == typeof(DateTime))
            {
                result = DateTime.Parse(text, culture, DateTimeStyles.None);
                return true;
            }

            if (targetType.IsEnum)
            {
                result = Enum.Parse(targetType, text, ignoreCase: true);
                return true;
            }

            result = System.Convert.ChangeType(text, targetType, culture);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }
}
