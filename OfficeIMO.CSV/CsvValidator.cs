#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV;

internal static class CsvValidator
{
    public static List<CsvValidationError> Validate(CsvDocument document, CsvSchema schema)
    {
        var errors = new List<CsvValidationError>();

        foreach (var column in schema.Columns)
        {
            if (!document.TryGetColumnIndexInternal(column.Name, out _) && column.IsRequired)
            {
                errors.Add(new CsvValidationError(-1, column.Name, "Required column is missing."));
            }
        }

        var rowIndex = 0;
        foreach (var row in document.AsEnumerable())
        {
            foreach (var column in schema.Columns)
            {
                if (!document.TryGetColumnIndexInternal(column.Name, out var index))
                {
                    continue;
                }

                var value = row[index];

                if (value is null || (value is string s && string.IsNullOrEmpty(s)))
                {
                    if (column.DefaultValue is not null)
                    {
                        row[index] = column.DefaultValue;
                        value = column.DefaultValue;
                    }
                    else if (column.IsRequired)
                    {
                        errors.Add(new CsvValidationError(rowIndex, column.Name, "Value is required."));
                        continue;
                    }
                }

                if (!TryConvertValue(value, column, document.Culture, document.DateTimeFormats, out var convertedValue, out var error))
                {
                    errors.Add(new CsvValidationError(rowIndex, column.Name, error ?? "Invalid value."));
                    continue;
                }

                foreach (var validator in column.Validators)
                {
                    if (!validator.Predicate(convertedValue))
                    {
                        errors.Add(new CsvValidationError(rowIndex, column.Name, validator.Message));
                    }
                }
            }

            rowIndex++;
        }

        return errors;
    }

    private static bool TryConvertValue(
        object? value,
        CsvSchemaColumn column,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        out object? convertedValue,
        out string? error)
    {
        convertedValue = value;
        error = null;

        if (column.Converter is { } converter)
        {
            try
            {
                convertedValue = converter(value, culture);
            }
            catch (Exception ex) when (ex is not CsvException)
            {
                error = $"Custom converter failed: {ex.Message}";
                return false;
            }
        }

        if (column.DataType is null || convertedValue is null || column.DataType.IsInstanceOfType(convertedValue))
        {
            return true;
        }

        if (!CsvValueConverter.TryConvert(convertedValue, column.DataType, culture, dateTimeFormats, out convertedValue, out error))
        {
            return false;
        }

        return true;
    }
}
