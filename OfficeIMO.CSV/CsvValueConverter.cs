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
                if (int.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Int32.";
                return false;
            }

            if (targetType == typeof(long))
            {
                if (long.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Int64.";
                return false;
            }

            if (targetType == typeof(short))
            {
                if (short.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Int16.";
                return false;
            }

            if (targetType == typeof(byte))
            {
                if (byte.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Byte.";
                return false;
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
                if (double.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Double.";
                return false;
            }

            if (targetType == typeof(decimal))
            {
                if (decimal.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Decimal.";
                return false;
            }

            if (targetType == typeof(float))
            {
                if (float.TryParse(text, NumberStyles.Any, culture, out var parsed))
                {
                    result = parsed;
                    return true;
                }

                error = $"Cannot parse '{text}' as Single.";
                return false;
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
