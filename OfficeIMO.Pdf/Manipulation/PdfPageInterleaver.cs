using System.Collections.ObjectModel;
using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Controls how an interleave handles sources with different selected page counts.</summary>
public enum PdfInterleaveRemainderMode {
    /// <summary>Appends remaining pages from longer sources in round-robin order.</summary>
    Append,
    /// <summary>Rejects inputs whose selected page counts differ.</summary>
    Reject
}

/// <summary>One source participating in an interleaved page composition.</summary>
public sealed class PdfInterleaveSource {
    private readonly byte[] _pdf;

    /// <summary>Creates an interleave source from PDF bytes.</summary>
    public PdfInterleaveSource(byte[] pdf, string? name = null) {
        Guard.NotNull(pdf, nameof(pdf));
        _pdf = (byte[])pdf.Clone();
        Name = name;
    }

    /// <summary>Optional source name carried into the output-page mapping.</summary>
    public string? Name { get; }
    /// <summary>Optional parser, password, permission, and resource-budget settings.</summary>
    public PdfReadOptions? ReadOptions { get; set; }
    /// <summary>Optional ordered page selector. All pages are used when omitted.</summary>
    public PdfPageSelector? Pages { get; set; }
    /// <summary>Reverses the selected page order before interleaving.</summary>
    public bool Reverse { get; set; }

    internal byte[] GetBytes() => (byte[])_pdf.Clone();
}

/// <summary>Controls interleaved page composition and document-level merge policy.</summary>
public sealed class PdfInterleaveOptions {
    /// <summary>Zero-based source whose metadata and keep-primary structures own the output catalog.</summary>
    public int PrimarySourceIndex { get; set; }
    /// <summary>Behavior when selected source page counts differ.</summary>
    public PdfInterleaveRemainderMode RemainderMode { get; set; } = PdfInterleaveRemainderMode.Append;
    /// <summary>Document-level merge, annotation-flattening, and page-resize policy.</summary>
    public PdfMergeOptions MergeOptions { get; set; } = new PdfMergeOptions();
}

/// <summary>Maps one output page to its source document and page.</summary>
public sealed class PdfInterleavePageMapping {
    internal PdfInterleavePageMapping(int outputPageNumber, int sourceIndex, string sourceName, int sourcePageNumber) {
        OutputPageNumber = outputPageNumber;
        SourceIndex = sourceIndex;
        SourceName = sourceName;
        SourcePageNumber = sourcePageNumber;
    }

    /// <summary>One-based page number in the composed PDF.</summary>
    public int OutputPageNumber { get; }
    /// <summary>Zero-based source index.</summary>
    public int SourceIndex { get; }
    /// <summary>Stable source name.</summary>
    public string SourceName { get; }
    /// <summary>One-based page number in the original source.</summary>
    public int SourcePageNumber { get; }
}

/// <summary>Interleaved PDF output with page provenance and document-structure decisions.</summary>
public sealed class PdfInterleaveResult {
    private readonly byte[] _pdf;
    private readonly PdfReadOptions _readOptions;

    internal PdfInterleaveResult(byte[] pdf, IReadOnlyList<PdfInterleavePageMapping> pages, PdfMergeReport mergeReport, PdfReadOptions readOptions) {
        _pdf = (byte[])pdf.Clone();
        Pages = pages;
        MergeReport = mergeReport;
        _readOptions = readOptions;
    }

    /// <summary>Output-page provenance in final page order.</summary>
    public IReadOnlyList<PdfInterleavePageMapping> Pages { get; }
    /// <summary>Document-level structure decisions applied during composition.</summary>
    public PdfMergeReport MergeReport { get; }
    /// <summary>Returns an independent copy of the composed PDF.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the composed PDF through the public document API.</summary>
    public PdfDocument ToDocument(PdfReadOptions? readOptions = null) => PdfDocument.Open(_pdf, readOptions ?? _readOptions);
}

/// <summary>Creates alternating or round-robin page compositions from multiple PDFs.</summary>
public static class PdfPageInterleaver {
    /// <summary>Interleaves complete PDFs in round-robin order.</summary>
    public static PdfInterleaveResult Interleave(params byte[][] pdfs) {
        Guard.NotNull(pdfs, nameof(pdfs));
        return Interleave(
            pdfs.Select((pdf, index) => new PdfInterleaveSource(pdf, "source-" + (index + 1).ToString(CultureInfo.InvariantCulture))),
            null);
    }

    /// <summary>Interleaves selected source pages and returns page provenance and merge-policy evidence.</summary>
    public static PdfInterleaveResult Interleave(IEnumerable<PdfInterleaveSource> sources, PdfInterleaveOptions? options = null) {
        Guard.NotNull(sources, nameof(sources));
        return PdfMerger.Interleave(sources.ToArray(), options ?? new PdfInterleaveOptions());
    }
}

internal static partial class PdfMerger {
    internal static PdfInterleaveResult Interleave(IReadOnlyList<PdfInterleaveSource> sources, PdfInterleaveOptions options) {
        Guard.NotNull(sources, nameof(sources));
        Guard.NotNull(options, nameof(options));
        if (sources.Count < 2) throw new ArgumentException("At least two PDFs must be supplied for interleaving.", nameof(sources));
        if (sources.Any(static source => source is null)) throw new ArgumentException("Interleave sources cannot contain null entries.", nameof(sources));
        if (options.PrimarySourceIndex < 0 || options.PrimarySourceIndex >= sources.Count) throw new ArgumentOutOfRangeException(nameof(options), "Primary source index must refer to an interleave source.");
        if (options.RemainderMode != PdfInterleaveRemainderMode.Append && options.RemainderMode != PdfInterleaveRemainderMode.Reject) throw new ArgumentOutOfRangeException(nameof(options), "Unknown interleave remainder mode.");
        Guard.NotNull(options.MergeOptions, nameof(options));
        ValidateInterleaveMergePolicy(options.MergeOptions.Policy);

        var sourceBytes = new byte[sources.Count][];
        var sourceDocuments = new PdfReadDocument[sources.Count];
        var selectedPageNumbers = new int[sources.Count][];
        var selectedPageObjectNumbers = new int[sources.Count][];
        var sourceNames = new string[sources.Count];
        var sourceSecurity = new PdfDocumentSecurityInfo[sources.Count];
        var sourcePermissionPolicy = new PdfPermissionPolicy[sources.Count];

        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            PdfInterleaveSource input = sources[sourceIndex];
            byte[] original = input.GetBytes();
            (PdfMutationPlan plan, PdfReadDocument plannedDocument) = PdfMutationPlanner.RequireFullRewriteDocument(
                original,
                PdfMutationOperation.MergeDocuments,
                input.ReadOptions);
            sourceSecurity[sourceIndex] = plan.Preflight.Probe.Security;
            sourcePermissionPolicy[sourceIndex] = plan.Preflight.PermissionPolicy;
            IReadOnlyList<int> resolved = input.Pages?.Resolve(plannedDocument.Pages.Count) ??
                Enumerable.Range(1, plannedDocument.Pages.Count).ToArray();
            int[] pageNumbers = resolved.ToArray();
            if (pageNumbers.Length == 0) throw new ArgumentException("Interleave source " + sourceIndex.ToString(CultureInfo.InvariantCulture) + " selected no pages.", nameof(sources));
            if (pageNumbers.Distinct().Count() != pageNumbers.Length) throw new ArgumentException("Interleave page selections cannot contain duplicate pages.", nameof(sources));
            if (input.Reverse) Array.Reverse(pageNumbers);

            byte[] prepared = PrepareMergeSource(original, options.MergeOptions, input.ReadOptions);
            PdfReadDocument preparedDocument = ReferenceEquals(prepared, original)
                ? plannedDocument
                : PdfReadDocument.Open(prepared, input.ReadOptions);
            sourceBytes[sourceIndex] = prepared;
            sourceDocuments[sourceIndex] = preparedDocument;
            selectedPageNumbers[sourceIndex] = pageNumbers;
            selectedPageObjectNumbers[sourceIndex] = pageNumbers.Select(page => preparedDocument.Pages[page - 1].ObjectNumber).ToArray();
            sourceNames[sourceIndex] = string.IsNullOrWhiteSpace(input.Name)
                ? "source-" + (sourceIndex + 1).ToString(CultureInfo.InvariantCulture)
                : input.Name!;
        }

        if (options.RemainderMode == PdfInterleaveRemainderMode.Reject && selectedPageNumbers.Select(static pages => pages.Length).Distinct().Count() != 1) {
            throw new InvalidOperationException("Interleave sources have different selected page counts and the remainder policy is Reject.");
        }

        var mappings = new List<PdfInterleavePageMapping>(selectedPageNumbers.Sum(static pages => pages.Length));
        var order = new List<(int SourceIndex, int PageObjectNumber)>();
        int maximumPages = selectedPageNumbers.Max(static pages => pages.Length);
        for (int pageIndex = 0; pageIndex < maximumPages; pageIndex++) {
            for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
                if (pageIndex >= selectedPageNumbers[sourceIndex].Length) continue;
                order.Add((sourceIndex, selectedPageObjectNumbers[sourceIndex][pageIndex]));
                mappings.Add(new PdfInterleavePageMapping(
                    mappings.Count + 1,
                    sourceIndex,
                    sourceNames[sourceIndex],
                    selectedPageNumbers[sourceIndex][pageIndex]));
            }
        }

        var outputIndexMaps = new Dictionary<int, int>[sources.Count];
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) outputIndexMaps[sourceIndex] = new Dictionary<int, int>();
        for (int outputIndex = 0; outputIndex < order.Count; outputIndex++) {
            outputIndexMaps[order[outputIndex].SourceIndex][order[outputIndex].PageObjectNumber] = outputIndex;
        }

        var importedSources = new ImportedSource[sources.Count];
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            importedSources[sourceIndex] = ImportSource(
                sourceBytes[sourceIndex],
                sourceIndex,
                selectedPageObjectNumbers[sourceIndex],
                mergedPageOffset: 0,
                outputIndexMaps[sourceIndex],
                sources[sourceIndex].ReadOptions,
                sourceDocuments[sourceIndex],
                sourceSecurity[sourceIndex],
                sourcePermissionPolicy[sourceIndex]);
        }

        byte[] composed = WriteMerged(
            importedSources,
            options.PrimarySourceIndex,
            order.Select(static page => new OutputPageReference(page.SourceIndex, page.PageObjectNumber)).ToArray());
        PdfReadOptions outputReadOptions = PdfReadOptions.WithMinimumInputBytes(
            sources[options.PrimarySourceIndex].ReadOptions,
            composed.LongLength);
        PdfMergeResult mergeResult = ApplyMergePolicy(
            composed,
            importedSources,
            options.PrimarySourceIndex,
            options.MergeOptions,
            outputReadOptions,
            order.Select(static page => page.SourceIndex).ToArray());
        byte[] output = mergeResult.ToBytes();
        PdfReadDocument reopened = PdfReadDocument.Open(output, PdfReadOptions.WithMinimumInputBytes(outputReadOptions, output.LongLength));
        if (reopened.Pages.Count != mappings.Count) throw new InvalidOperationException("Interleaved PDF page count does not match its provenance report.");
        return new PdfInterleaveResult(
            output,
            new ReadOnlyCollection<PdfInterleavePageMapping>(mappings),
            mergeResult.Report,
            mergeResult.ReadOptions);
    }

    private static void ValidateInterleaveMergePolicy(PdfMergePolicy policy) {
        Guard.NotNull(policy, nameof(policy));
        if (policy.Outlines == PdfMergeStructureMode.Combine ||
            policy.NamedDestinations == PdfMergeStructureMode.Combine ||
            policy.PageLabels == PdfMergeStructureMode.Combine ||
            policy.Forms == PdfMergeStructureMode.Combine) {
            throw new NotSupportedException("Interleaving does not combine page-addressed outlines, destinations, labels, or form trees. Use KeepPrimary, Drop, or RejectIncoming so page ownership stays explicit.");
        }
    }
}
