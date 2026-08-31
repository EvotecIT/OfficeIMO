#nullable enable

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Data;

internal sealed class ExplicitRowMappingPlan<T> where T : new() {
    private readonly MappingBinding[] _bindings;

    private ExplicitRowMappingPlan(MappingBinding[] bindings) {
        _bindings = bindings;
    }

    internal bool IsEmpty => _bindings.Length == 0;

    internal static ExplicitRowMappingPlan<T> Create(
        IReadOnlyList<string> headers,
        Action<RowMapper<T>> configure) {
        var mapper = new RowMapper<T>();
        configure(mapper);
        MappingBinding[] bindings = mapper.Entries
            .Select(entry => new MappingBinding(FindColumn(headers, entry.ColumnNames), entry))
            .ToArray();
        return new ExplicitRowMappingPlan<T>(bindings);
    }

    internal T MapRow(
        Func<int, object?> getValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Entry.Apply(
                instance,
                getValue(binding.ColumnIndex),
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
        return instance;
    }

    internal T MapValues(
        object?[] values,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Entry.Apply(
                instance,
                values[binding.ColumnIndex],
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
        return instance;
    }

    internal T MapReaderRow(
        DbDataReader reader,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Entry.ApplyReader(
                instance,
                reader,
                binding.ColumnIndex,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
        return instance;
    }

    private static int FindColumn(IReadOnlyList<string> headers, IReadOnlyList<string> columnNames) {
        for (int nameIndex = 0; nameIndex < columnNames.Count; nameIndex++) {
            int found = -1;
            for (int headerIndex = 0; headerIndex < headers.Count; headerIndex++) {
                if (!string.Equals(
                        headers[headerIndex],
                        columnNames[nameIndex],
                        StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (found >= 0) {
                    throw new DataMappingException(
                        $"Column name '{columnNames[nameIndex]}' matches multiple source columns.");
                }
                found = headerIndex;
            }

            if (found >= 0) return found;
        }

        throw new DataMappingException(
            $"None of the columns '{string.Join("', '", columnNames)}' were found.");
    }

    private readonly struct MappingBinding {
        internal MappingBinding(int columnIndex, IRowMappingEntry<T> entry) {
            ColumnIndex = columnIndex;
            Entry = entry;
        }

        internal int ColumnIndex { get; }
        internal IRowMappingEntry<T> Entry { get; }
    }
}
