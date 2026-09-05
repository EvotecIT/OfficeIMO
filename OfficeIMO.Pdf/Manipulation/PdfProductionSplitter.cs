using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

/// <summary>Reason a production split part ended at its final source page.</summary>
public enum PdfProductionSplitReason {
    /// <summary>The source document ended.</summary>
    EndOfDocument,
    /// <summary>The configured maximum page count was reached.</summary>
    PageCount,
    /// <summary>The next page matched the configured content boundary.</summary>
    ContentBoundary,
    /// <summary>Adding the next page would exceed the target artifact size.</summary>
    TargetSize
}

/// <summary>Controls page-count, content-boundary, and target-size PDF splitting.</summary>
public sealed class PdfProductionSplitOptions {
    /// <summary>Optional maximum number of source pages in one part.</summary>
    public int? MaximumPagesPerPart { get; set; }

    /// <summary>Optional target artifact size in bytes. A single source page may exceed the target.</summary>
    public long? TargetPartSizeBytes { get; set; }

    /// <summary>Optional text marker that starts a new part when found on a page.</summary>
    public string? BoundaryText { get; set; }

    /// <summary>Uses ordinal case-insensitive matching for <see cref="BoundaryText"/>.</summary>
    public bool IgnoreBoundaryTextCase { get; set; } = true;

    /// <summary>Maximum generated candidate artifacts used while enforcing target size.</summary>
    public int MaximumArtifactProbes { get; set; } = 100_000;

    /// <summary>Maximum cumulative bytes generated across candidate and final artifacts.</summary>
    public long MaximumCumulativeArtifactBytes { get; set; } = 4L * 1024L * 1024L * 1024L;
}

/// <summary>One generated production split artifact.</summary>
public sealed class PdfProductionSplitPart {
    private readonly byte[] _pdf;
    private readonly PdfLoadOptions _readOptions;

    internal PdfProductionSplitPart(
        int partNumber,
        byte[] pdf,
        IReadOnlyList<int> sourcePages,
        PdfProductionSplitReason reason,
        bool exceedsTargetSize,
        PdfLoadOptions readOptions) {
        PartNumber = partNumber;
        _pdf = (byte[])pdf.Clone();
        SourcePages = sourcePages;
        Reason = reason;
        ExceedsTargetSize = exceedsTargetSize;
        _readOptions = readOptions;
    }

    /// <summary>One-based part number.</summary>
    public int PartNumber { get; }
    /// <summary>One-based source pages contained in this part.</summary>
    public IReadOnlyList<int> SourcePages { get; }
    /// <summary>Reason this part ended at its final source page.</summary>
    public PdfProductionSplitReason Reason { get; }
    /// <summary>True when this part exceeds the target because one source page cannot be split further.</summary>
    public bool ExceedsTargetSize { get; }
    /// <summary>Generated artifact size in bytes.</summary>
    public long SizeBytes => _pdf.LongLength;
    /// <summary>Returns an independent copy of the generated part.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the generated part through the public document API.</summary>
    public PdfDocument ToDocument(PdfLoadOptions? readOptions = null) => PdfDocument.Load(_pdf, readOptions ?? _readOptions);
}

/// <summary>Generated production split artifacts and bounded-probe evidence.</summary>
public sealed class PdfProductionSplitResult : IOfficeResult<IReadOnlyList<PdfProductionSplitPart>> {
    internal PdfProductionSplitResult(int sourcePageCount, IReadOnlyList<PdfProductionSplitPart> parts, int artifactProbeCount, long cumulativeArtifactBytes) {
        SourcePageCount = sourcePageCount;
        Parts = parts;
        ArtifactProbeCount = artifactProbeCount;
        CumulativeArtifactBytes = cumulativeArtifactBytes;
    }

    /// <summary>Page count in the source PDF.</summary>
    public int SourcePageCount { get; }
    /// <summary>Generated artifacts in source-page order.</summary>
    public IReadOnlyList<PdfProductionSplitPart> Parts { get; }

    /// <inheritdoc />
    public bool Succeeded => true;

    /// <inheritdoc />
    public IReadOnlyList<PdfProductionSplitPart> Value => Parts;

    /// <inheritdoc />
    public IReadOnlyList<PdfProductionSplitPart> RequireValue() => Parts;
    /// <summary>Number of candidate artifacts generated to satisfy the split policy.</summary>
    public int ArtifactProbeCount { get; }
    /// <summary>Cumulative candidate and final artifact bytes generated while planning the split.</summary>
    public long CumulativeArtifactBytes { get; }
}

/// <summary>Creates bounded, report-driven PDF production splits.</summary>
internal static class PdfProductionSplitter {
    /// <summary>Splits a PDF using page-count, text-boundary, and target-size policies.</summary>
    public static PdfProductionSplitResult Split(byte[] pdf, PdfProductionSplitOptions options, PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(options, nameof(options));
        ValidateOptions(options);
        (_, PdfReadDocument document) = PdfMutationPlanner.RequireFullRewriteDocument(pdf, PdfMutationOperation.ExtractPages, readOptions);
        if (document.Pages.Count == 0) throw new ArgumentException("PDF does not contain any pages.", nameof(pdf));

        List<PlannedPart> planned = PlanStructuralBoundaries(document, options);
        var output = new List<PdfProductionSplitPart>();
        int probeCount = 0;
        long cumulativeArtifactBytes = 0L;
        foreach (PlannedPart group in planned) {
            if (!options.TargetPartSizeBytes.HasValue) {
                byte[] artifact = Extract(pdf, group.Pages, readOptions, options, ref probeCount, ref cumulativeArtifactBytes);
                AddVerifiedPart(output, pdf, artifact, group.Pages, group.Reason, exceedsTarget: false, readOptions);
                continue;
            }

            SplitGroupByTargetSize(pdf, group, options, readOptions, output, ref probeCount, ref cumulativeArtifactBytes);
        }

        int[] emittedPages = output.SelectMany(static part => part.SourcePages).ToArray();
        int[] expectedPages = Enumerable.Range(1, document.Pages.Count).ToArray();
        if (!emittedPages.SequenceEqual(expectedPages)) throw new InvalidOperationException("Production split page coverage does not match the source document.");
        return new PdfProductionSplitResult(
            document.Pages.Count,
            new ReadOnlyCollection<PdfProductionSplitPart>(output),
            probeCount,
            cumulativeArtifactBytes);
    }

    private static List<PlannedPart> PlanStructuralBoundaries(PdfReadDocument document, PdfProductionSplitOptions options) {
        var result = new List<PlannedPart>();
        var current = new List<int>();
        StringComparison comparison = options.IgnoreBoundaryTextCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        for (int pageNumber = 1; pageNumber <= document.Pages.Count; pageNumber++) {
            bool beginsAtContentBoundary = !string.IsNullOrEmpty(options.BoundaryText) &&
                Contains(document.Pages[pageNumber - 1].ExtractText(), options.BoundaryText!, comparison);
            if (beginsAtContentBoundary && current.Count > 0) {
                result.Add(new PlannedPart(current.ToArray(), PdfProductionSplitReason.ContentBoundary));
                current.Clear();
            }

            current.Add(pageNumber);
            if (options.MaximumPagesPerPart.HasValue && current.Count == options.MaximumPagesPerPart.Value) {
                result.Add(new PlannedPart(current.ToArray(), PdfProductionSplitReason.PageCount));
                current.Clear();
            }
        }

        if (current.Count > 0) result.Add(new PlannedPart(current.ToArray(), PdfProductionSplitReason.EndOfDocument));
        if (result.Count > 0 && result[result.Count - 1].Pages[result[result.Count - 1].Pages.Length - 1] == document.Pages.Count) {
            PlannedPart last = result[result.Count - 1];
            result[result.Count - 1] = new PlannedPart(last.Pages, PdfProductionSplitReason.EndOfDocument);
        }
        return result;
    }

    private static void SplitGroupByTargetSize(
        byte[] source,
        PlannedPart group,
        PdfProductionSplitOptions options,
        PdfLoadOptions? readOptions,
        List<PdfProductionSplitPart> output,
        ref int probeCount,
        ref long cumulativeArtifactBytes) {
        long targetBytes = options.TargetPartSizeBytes!.Value;
        int startIndex = 0;
        while (startIndex < group.Pages.Length) {
            int remainingPages = group.Pages.Length - startIndex;
            int bestPageCount = 1;
            int[] bestPages = SlicePages(group.Pages, startIndex, bestPageCount);
            byte[] bestArtifact = Extract(source, bestPages, readOptions, options, ref probeCount, ref cumulativeArtifactBytes);
            if (bestArtifact.LongLength > targetBytes) {
                PdfProductionSplitReason reason = remainingPages == 1 ? group.Reason : PdfProductionSplitReason.TargetSize;
                AddVerifiedPart(output, source, bestArtifact, bestPages, reason, exceedsTarget: true, readOptions);
                startIndex++;
                continue;
            }

            int firstOverflowPageCount = 0;
            int candidatePageCount = 1;
            while (candidatePageCount < remainingPages) {
                int nextPageCount = candidatePageCount > remainingPages / 2
                    ? remainingPages
                    : candidatePageCount * 2;
                int[] candidatePages = SlicePages(group.Pages, startIndex, nextPageCount);
                byte[] candidateArtifact = Extract(source, candidatePages, readOptions, options, ref probeCount, ref cumulativeArtifactBytes);
                if (candidateArtifact.LongLength > targetBytes) {
                    firstOverflowPageCount = nextPageCount;
                    break;
                }

                bestPageCount = nextPageCount;
                bestPages = candidatePages;
                bestArtifact = candidateArtifact;
                candidatePageCount = nextPageCount;
            }

            int low = bestPageCount + 1;
            int high = firstOverflowPageCount > 0 ? firstOverflowPageCount - 1 : bestPageCount;
            while (low <= high) {
                int middle = low + ((high - low) / 2);
                int[] candidatePages = SlicePages(group.Pages, startIndex, middle);
                byte[] candidateArtifact = Extract(source, candidatePages, readOptions, options, ref probeCount, ref cumulativeArtifactBytes);
                if (candidateArtifact.LongLength <= targetBytes) {
                    bestPageCount = middle;
                    bestPages = candidatePages;
                    bestArtifact = candidateArtifact;
                    low = middle + 1;
                } else {
                    high = middle - 1;
                }
            }

            bool reachesGroupEnd = startIndex + bestPageCount == group.Pages.Length;
            AddVerifiedPart(
                output,
                source,
                bestArtifact,
                bestPages,
                reachesGroupEnd ? group.Reason : PdfProductionSplitReason.TargetSize,
                exceedsTarget: false,
                readOptions);
            startIndex += bestPageCount;
        }
    }

    private static int[] SlicePages(int[] pages, int startIndex, int count) {
        var result = new int[count];
        Array.Copy(pages, startIndex, result, 0, count);
        return result;
    }

    private static byte[] Extract(
        byte[] source,
        IEnumerable<int> pages,
        PdfLoadOptions? readOptions,
        PdfProductionSplitOptions options,
        ref int probeCount,
        ref long cumulativeArtifactBytes) {
        probeCount = checked(probeCount + 1);
        if (probeCount > options.MaximumArtifactProbes) {
            throw new InvalidOperationException("Production splitting exceeded the configured artifact-probe budget.");
        }
        long remainingArtifactBytes = options.MaximumCumulativeArtifactBytes - cumulativeArtifactBytes;
        if (remainingArtifactBytes <= 0L) {
            throw new InvalidOperationException("Production splitting exceeded the configured cumulative artifact-byte budget.");
        }
        byte[] artifact;
        try {
            artifact = PdfPageExtractor.ExtractPages(source, pages, readOptions, remainingArtifactBytes);
        } catch (InvalidDataException exception) {
            throw new InvalidOperationException("Production splitting exceeded the configured cumulative artifact-byte budget.", exception);
        }
        cumulativeArtifactBytes = checked(cumulativeArtifactBytes + artifact.LongLength);
        if (cumulativeArtifactBytes > options.MaximumCumulativeArtifactBytes) {
            throw new InvalidOperationException("Production splitting exceeded the configured cumulative artifact-byte budget.");
        }
        return artifact;
    }

    private static void AddVerifiedPart(
        List<PdfProductionSplitPart> output,
        byte[] source,
        byte[] artifact,
        IEnumerable<int> pages,
        PdfProductionSplitReason reason,
        bool exceedsTarget,
        PdfLoadOptions? readOptions) {
        int[] pageNumbers = pages.ToArray();
        PdfLoadOptions strictOptions = CreateStrictReadOptions(readOptions, source, artifact);
        PdfReadDocument reopened = PdfReadDocument.Open(artifact, strictOptions);
        if (reopened.Pages.Count != pageNumbers.Length) throw new InvalidOperationException("Generated split artifact page count does not match its report.");
        output.Add(new PdfProductionSplitPart(
            output.Count + 1,
            artifact,
            Array.AsReadOnly(pageNumbers),
            reason,
            exceedsTarget,
            strictOptions));
    }

    private static PdfLoadOptions CreateStrictReadOptions(PdfLoadOptions? readOptions, byte[] source, byte[] artifact) {
        PdfLoadOptions effective = PdfLoadOptions.ForGeneratedOutput(readOptions, source, artifact);
        return new PdfLoadOptions {
            ParsingMode = PdfParsingMode.Strict,
            Limits = effective.Limits,
            Password = effective.Password,
            AesCryptographyProvider = effective.AesCryptographyProvider,
            PermissionPolicy = effective.PermissionPolicy,
            PreferToUnicode = effective.PreferToUnicode,
            UseWinAnsiFallback = effective.UseWinAnsiFallback,
            AdjustKerningFromTJ = effective.AdjustKerningFromTJ,
            IncludeArtifactText = effective.IncludeArtifactText
        };
    }

    private static void ValidateOptions(PdfProductionSplitOptions options) {
        if (options.MaximumPagesPerPart <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum pages per part must be positive.");
        if (options.TargetPartSizeBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(options), "Target part size must be positive.");
        if (options.BoundaryText is not null && options.BoundaryText.Length == 0) throw new ArgumentException("Boundary text cannot be empty.", nameof(options));
        if (options.BoundaryText is not null && options.BoundaryText.Length > 4096) throw new ArgumentException("Boundary text is too long.", nameof(options));
        if (!options.MaximumPagesPerPart.HasValue && !options.TargetPartSizeBytes.HasValue && options.BoundaryText is null) {
            throw new ArgumentException("At least one production split criterion must be configured.", nameof(options));
        }
        if (options.MaximumArtifactProbes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Maximum artifact probes must be positive.");
        if (options.MaximumCumulativeArtifactBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(options), "Maximum cumulative artifact bytes must be positive.");
    }

    private static bool Contains(string source, string value, StringComparison comparison) {
#if NET472 || NETSTANDARD2_0
        return source.IndexOf(value, comparison) >= 0;
#else
        return source.Contains(value, comparison);
#endif
    }

    private sealed class PlannedPart {
        internal PlannedPart(int[] pages, PdfProductionSplitReason reason) {
            Pages = pages;
            Reason = reason;
        }

        internal int[] Pages { get; }
        internal PdfProductionSplitReason Reason { get; }
    }
}
