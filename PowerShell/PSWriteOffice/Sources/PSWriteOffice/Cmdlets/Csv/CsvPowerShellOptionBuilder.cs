using System;
using System.Collections;
using System.Collections.Generic;
using System.IO.Compression;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal static class CsvPowerShellOptionBuilder
{
    public static void ApplyLoadOptions(
        CsvLoadOptions options,
        CsvDuplicateHeaderBehavior duplicateHeaderBehavior,
        string? nullValue,
        string[]? dateTimeFormats,
        CsvQuoteParsingMode quoteParsingMode,
        IDictionary? staticColumns,
        CsvCompressionType compressionType,
        long? maxDecompressedBytes)
    {
        options.DuplicateHeaderBehavior = duplicateHeaderBehavior;
        options.NullValue = nullValue;
        options.DateTimeFormats = dateTimeFormats;
        options.QuoteParsingMode = quoteParsingMode;
        options.StaticColumns = ToStaticColumnDictionary(staticColumns);
        options.CompressionType = compressionType;
        options.MaxDecompressedBytes = maxDecompressedBytes;
    }

    public static void ApplyTextLoadOptions(
        CsvLoadOptions options,
        CsvDuplicateHeaderBehavior duplicateHeaderBehavior,
        string? nullValue,
        string[]? dateTimeFormats,
        CsvQuoteParsingMode quoteParsingMode,
        IDictionary? staticColumns)
    {
        options.DuplicateHeaderBehavior = duplicateHeaderBehavior;
        options.NullValue = nullValue;
        options.DateTimeFormats = dateTimeFormats;
        options.QuoteParsingMode = quoteParsingMode;
        options.StaticColumns = ToStaticColumnDictionary(staticColumns);
    }

    public static void ApplySaveOptions(
        CsvSaveOptions options,
        string? nullValue,
        string? dateTimeFormat,
        bool useUtc,
        CsvCompressionType compressionType,
        CompressionLevel compressionLevel)
    {
        options.NullValue = nullValue;
        options.DateTimeFormat = dateTimeFormat;
        options.UseUtc = useUtc;
        options.CompressionType = compressionType;
        options.CompressionLevel = compressionLevel;
    }

    private static IReadOnlyDictionary<string, object?>? ToStaticColumnDictionary(IDictionary? source)
    {
        if (source == null || source.Count == 0)
        {
            return null;
        }

        var result = new Dictionary<string, object?>(source.Count, StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry entry in source)
        {
            if (entry.Key == null)
            {
                throw new ArgumentException("Static column names cannot be null.", nameof(source));
            }

            var name = entry.Key.ToString();
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("Static column names cannot be empty.", nameof(source));
            }

            result[name] = entry.Value;
        }

        return result;
    }
}
