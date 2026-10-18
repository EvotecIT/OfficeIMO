using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Email;

namespace PSWriteOffice.Services.Reader;

internal static class ReaderCommandUtilities
{
    private static readonly Lazy<OfficeDocumentReader> SharedReader = new(() => CreateReader());

    internal static OfficeDocumentReader Reader => SharedReader.Value;

    internal static string ResolvePath(PSCmdlet cmdlet, string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("Path cannot be empty.", nameof(path));
        }

        var providerPath = cmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(cmdlet.SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    internal static IReadOnlyList<string> ResolvePaths(PSCmdlet cmdlet, string path)
    {
        try
        {
            var resolved = cmdlet.SessionState.Path.GetResolvedProviderPathFromPSPath(path, out _);
            return resolved.Count > 0
                ? resolved.ToArray()
                : new[] { ResolvePath(cmdlet, path) };
        }
        catch (ItemNotFoundException)
        {
            return new[] { ResolvePath(cmdlet, path) };
        }
    }

    internal static OfficeDocumentReaderBuilder CreateBuilder(ReaderAllOptions? options = null)
    {
        return new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers(options);
    }

    internal static OfficeDocumentReader CreateReader(
        ReaderAllOptions? options = null,
        int? maxConcurrentReads = null)
    {
        OfficeDocumentReaderBuilder builder = CreateBuilder(options);
        if (maxConcurrentReads.HasValue)
        {
            builder.WithMaxConcurrentReads(maxConcurrentReads.Value);
        }

        return builder.Build();
    }

    internal static void ValidateBatchConcurrency(
        OfficeDocumentReader reader,
        ReaderBatchOptions batchOptions)
    {
        if (batchOptions.MaxDegreeOfParallelism.HasValue &&
            batchOptions.MaxDegreeOfParallelism.Value > reader.MaxConcurrentReads)
        {
            throw new PSArgumentException(
                $"The requested batch concurrency ({batchOptions.MaxDegreeOfParallelism.Value}) exceeds the " +
                $"immutable Reader limit ({reader.MaxConcurrentReads}). Create one with " +
                "New-OfficeDocumentReader -MaxConcurrentReads set to the same or a higher value.");
        }
    }

    internal static ReaderCommandConfiguration BuildReadConfiguration(
        long? maxInputBytes,
        long? openXmlMaxCharactersInPart,
        int? maxChars,
        int? maxTableRows,
        bool includeWordFootnotes,
        bool includePowerPointNotes,
        bool excelHeadersInFirstRow,
        int? excelChunkRows,
        string? excelSheetName,
        string? excelA1Range,
        bool markdownChunkByHeadings,
        bool computeHashes,
        bool includePageLocations = false,
        int? maxStoreItems = null)
    {
        var readerOptions = new ReaderOptions
        {
            ComputeHashes = computeHashes
        };

        if (maxInputBytes.HasValue)
        {
            readerOptions.MaxInputBytes = maxInputBytes.Value;
        }

        if (openXmlMaxCharactersInPart.HasValue)
        {
            readerOptions.OpenXmlMaxCharactersInPart = openXmlMaxCharactersInPart.Value;
        }

        if (maxChars.HasValue)
        {
            readerOptions.MaxChars = maxChars.Value;
        }

        if (maxTableRows.HasValue)
        {
            readerOptions.MaxTableRows = maxTableRows.Value;
        }

        var hasHandlerOverrides =
            !includeWordFootnotes ||
            !includePowerPointNotes ||
            !excelHeadersInFirstRow ||
            excelChunkRows.HasValue ||
            !string.IsNullOrWhiteSpace(excelSheetName) ||
            !string.IsNullOrWhiteSpace(excelA1Range) ||
            !markdownChunkByHeadings ||
            includePageLocations ||
            maxStoreItems.HasValue;
        ReaderAllOptions? handlerOptions = hasHandlerOverrides
            ? new ReaderAllOptions
            {
                Word = new OfficeIMO.Reader.Word.ReaderWordOptions
                {
                    IncludeFootnotes = includeWordFootnotes,
                    IncludePageLocations = includePageLocations
                },
                PowerPoint = new OfficeIMO.Reader.PowerPoint.ReaderPowerPointOptions
                {
                    IncludeNotes = includePowerPointNotes
                },
                Excel = new OfficeIMO.Reader.Excel.ReaderExcelOptions
                {
                    HeadersInFirstRow = excelHeadersInFirstRow,
                    ChunkRows = excelChunkRows ?? 200,
                    SheetName = string.IsNullOrWhiteSpace(excelSheetName) ? null : excelSheetName,
                    A1Range = string.IsNullOrWhiteSpace(excelA1Range) ? null : excelA1Range
                },
                Markdown = new OfficeIMO.Reader.Markdown.ReaderMarkdownOptions
                {
                    ChunkByHeadings = markdownChunkByHeadings
                },
                Rtf = new OfficeIMO.Reader.Rtf.ReaderRtfOptions
                {
                    IncludePageLocations = includePageLocations
                },
                Email = maxStoreItems.HasValue
                    ? new ReaderEmailHandlersOptions
                    {
                        Stores = new ReaderEmailStoreOptions
                        {
                            MaxItems = maxStoreItems.Value
                        }
                    }
                    : null
            }
            : null;

        return new ReaderCommandConfiguration(readerOptions, handlerOptions);
    }

    internal static ReaderCommandConfiguration BuildSearchConfiguration(
        bool includePageLocations,
        int? maxStoreItems)
    {
        return BuildReadConfiguration(
            maxInputBytes: null,
            openXmlMaxCharactersInPart: null,
            maxChars: null,
            maxTableRows: null,
            includeWordFootnotes: true,
            includePowerPointNotes: true,
            excelHeadersInFirstRow: true,
            excelChunkRows: null,
            excelSheetName: null,
            excelA1Range: null,
            markdownChunkByHeadings: true,
            computeHashes: false,
            includePageLocations,
            maxStoreItems);
    }

    internal static bool HasSourceLimit(OfficeDocumentReadResult document)
    {
        return document.Metadata.Any(static entry =>
                   string.Equals(entry.Name, "SelectionLimitReached", StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(entry.Value, "True", StringComparison.OrdinalIgnoreCase)) ||
               document.Diagnostics.Any(static diagnostic =>
                   diagnostic.Category == OfficeDocumentDiagnosticCategory.Limit);
    }

    internal static IReadOnlyList<string> CollectDocumentPaths(
        IEnumerable<string> paths,
        int maxDocuments,
        IEnumerable<string>? extensions,
        out bool limitReached)
    {
        if (paths == null) throw new ArgumentNullException(nameof(paths));
        if (maxDocuments < 1) throw new ArgumentOutOfRangeException(nameof(maxDocuments));

        var comparer = Path.DirectorySeparatorChar == '\\'
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;
        var unique = new HashSet<string>(comparer);
        HashSet<string>? allowedExtensions = extensions == null
            ? null
            : new HashSet<string>(extensions, StringComparer.OrdinalIgnoreCase);
        var result = new List<string>(Math.Min(maxDocuments, 256));
        limitReached = false;
        foreach (var path in paths)
        {
            if (allowedExtensions is { Count: > 0 } &&
                !allowedExtensions.Contains(Path.GetExtension(path)))
            {
                continue;
            }

            if (!unique.Add(path))
            {
                continue;
            }
            if (result.Count >= maxDocuments)
            {
                limitReached = true;
                break;
            }

            result.Add(path);
        }

        return result;
    }

    internal static ReaderFolderOptions BuildFolderOptions(bool recurse, int? maxFiles, long? maxTotalBytes, string[]? extension)
    {
        var options = new ReaderFolderOptions
        {
            Recurse = recurse
        };

        if (maxFiles.HasValue)
        {
            options.MaxFiles = maxFiles.Value;
        }

        if (maxTotalBytes.HasValue)
        {
            options.MaxTotalBytes = maxTotalBytes.Value;
        }

        if (extension != null && extension.Length > 0)
        {
            options.Extensions = extension
                .Where(static value => !string.IsNullOrWhiteSpace(value))
                .Select(static value => value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value)
                .ToArray();
        }

        return options;
    }
}

internal sealed class ReaderCommandConfiguration
{
    internal ReaderCommandConfiguration(ReaderOptions readerOptions, ReaderAllOptions? handlerOptions)
    {
        ReaderOptions = readerOptions;
        HandlerOptions = handlerOptions;
    }

    internal ReaderOptions ReaderOptions { get; }

    internal ReaderAllOptions? HandlerOptions { get; }
}
