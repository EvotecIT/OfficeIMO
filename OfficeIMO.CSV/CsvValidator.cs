#nullable enable

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

                if (column.DataType is not null)
                {
                    if (!CsvValueConverter.TryConvert(value, column.DataType, document.Culture, out _, out var error))
                    {
                        errors.Add(new CsvValidationError(rowIndex, column.Name, error ?? "Invalid value."));
                        continue;
                    }
                }

                foreach (var validator in column.Validators)
                {
                    if (!validator.Predicate(value))
                    {
                        errors.Add(new CsvValidationError(rowIndex, column.Name, validator.Message));
                    }
                }
            }

            rowIndex++;
        }

        return errors;
    }
}
