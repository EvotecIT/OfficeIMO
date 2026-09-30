#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Data;

/// <summary>Controls one explicit column binding in <see cref="RowMapper{T}"/>.</summary>
public sealed class RowMappingColumnOptions {
    /// <summary>
    /// Gets or sets whether a missing source column leaves the model's initialized value
    /// unchanged. Present columns still convert and assign normally. Default is false.
    /// </summary>
    public bool Optional { get; set; }

    /// <summary>Gets or sets a culture overriding the reader's conversion culture for this binding.</summary>
    public CultureInfo? Culture { get; set; }

    /// <summary>
    /// Gets or sets additional exact date/time formats overriding the reader's format list.
    /// The shared conversion pipeline tries these before ordinary culture-based parsing.
    /// </summary>
    public IReadOnlyList<string>? DateTimeFormats { get; set; }

    /// <summary>
    /// Gets or sets a converter overriding the reader's converter for this binding.
    /// Returning false selects built-in conversion; handled results, including null, take
    /// precedence. Parallel mapping may invoke the delegate concurrently.
    /// </summary>
    public Func<object, Type, CultureInfo, (bool ok, object? value)>? TypeConverter { get; set; }

    internal RowMappingColumnOptions Snapshot() {
        if (DateTimeFormats is not null && DateTimeFormats.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Date/time formats cannot contain empty values.", nameof(DateTimeFormats));
        return new RowMappingColumnOptions {
            Optional = Optional,
            Culture = Culture is null ? null : CultureInfo.ReadOnly((CultureInfo)Culture.Clone()),
            DateTimeFormats = DateTimeFormats?.ToArray(),
            TypeConverter = TypeConverter
        };
    }
}
