using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Reader;

/// <summary>
/// Describes how a page-like location was obtained from the source format.
/// </summary>
public enum OfficeDocumentPageProvenance {
    /// <summary>The producer did not identify how the page was obtained.</summary>
    Unknown = 0,

    /// <summary>The source format stores fixed physical pages.</summary>
    Native = 1,

    /// <summary>The page was calculated by an OfficeIMO layout engine.</summary>
    Computed = 2,

    /// <summary>The page was reconstructed from explicit or saved page-break hints.</summary>
    ExplicitBreak = 3,

    /// <summary>The source stores a page-like logical container such as a slide or sheet.</summary>
    LogicalContainer = 4
}

/// <summary>
/// Page-like location of one normalized source block.
/// </summary>
public sealed class OfficeDocumentPageLocation {
    internal OfficeDocumentPageLocation(
        OfficeDocumentPage page,
        int totalPageCount,
        OfficeDocumentPageProvenance provenance,
        IReadOnlyList<OfficeDocumentRegion> regions) {
        Page = page;
        TotalPageCount = totalPageCount;
        Provenance = provenance;
        Regions = regions;
    }

    /// <summary>The page-like container holding the source block.</summary>
    public OfficeDocumentPage Page { get; }

    /// <summary>One-based physical page number when available.</summary>
    public int? Number => Page.Number;

    /// <summary>Total physical page count known for the read operation.</summary>
    public int TotalPageCount { get; }

    /// <summary>How the page location was obtained.</summary>
    public OfficeDocumentPageProvenance Provenance { get; }

    /// <summary>Page-specific regions occupied by the source block.</summary>
    public IReadOnlyList<OfficeDocumentRegion> Regions { get; }

    /// <summary>Human-readable page label suitable for search results and citations.</summary>
    public string Display {
        get {
            string label = !string.IsNullOrWhiteSpace(Page.Name)
                ? Page.Name!
                : Number.HasValue
                    ? "Page " + Number.Value.ToString(CultureInfo.InvariantCulture)
                    : "Page";
            return TotalPageCount > 0 && Number.HasValue
                ? label + " of " + TotalPageCount.ToString(CultureInfo.InvariantCulture)
                : label;
        }
    }
}

/// <summary>
/// Location and search helpers for the normalized document read result.
/// </summary>
public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>
    /// Locates a normalized block in every page-like container that references it.
    /// </summary>
    public static IReadOnlyList<OfficeDocumentPageLocation> Locate(
        this OfficeDocumentReadResult document,
        OfficeDocumentBlock block) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (block == null) throw new ArgumentNullException(nameof(block));
        return Locate(document, block.Id, block);
    }

    /// <summary>
    /// Locates a normalized block by its stable identifier.
    /// </summary>
    public static IReadOnlyList<OfficeDocumentPageLocation> Locate(
        this OfficeDocumentReadResult document,
        string blockId) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(blockId)) throw new ArgumentException("Block identifier cannot be empty.", nameof(blockId));
        return Locate(document, blockId, fallbackBlock: null);
    }

    /// <summary>
    /// Returns the best available total physical page count for this read result.
    /// </summary>
    public static int GetTotalPageCount(this OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        int diagnosticPageCount = 0;
        foreach (ReaderChunk chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) {
            if (chunk.Diagnostics != null) {
                diagnosticPageCount = Math.Max(diagnosticPageCount, chunk.Diagnostics.PageCount);
            }
        }

        int metadataPageCount = 0;
        foreach (OfficeDocumentMetadataEntry entry in document.Metadata ?? Array.Empty<OfficeDocumentMetadataEntry>()) {
            if ((string.Equals(entry.Name, "PageCount", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(entry.Name, "NumberOfPages", StringComparison.OrdinalIgnoreCase)) &&
                int.TryParse(
                    entry.Value,
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out int parsedPageCount) &&
                parsedPageCount > 0) {
                metadataPageCount = Math.Max(metadataPageCount, parsedPageCount);
            }
        }

        int highestNumber = 0;
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            highestNumber = Math.Max(highestNumber, page.Number ?? 0);
        }
        return Math.Max(
            diagnosticPageCount,
            Math.Max(metadataPageCount, Math.Max(highestNumber, document.Pages?.Count ?? 0)));
    }

    /// <summary>
    /// Returns the page provenance advertised by the producing adapter.
    /// </summary>
    public static OfficeDocumentPageProvenance GetPageProvenance(this OfficeDocumentReadResult document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IReadOnlyList<string> capabilities = document.CapabilitiesUsed ?? Array.Empty<string>();
        if (capabilities.Contains("officeimo.reader.pdf.pages.native", StringComparer.Ordinal)) {
            return OfficeDocumentPageProvenance.Native;
        }
        if (capabilities.Contains("officeimo.reader.word.pages.computed", StringComparer.Ordinal)) {
            return OfficeDocumentPageProvenance.Computed;
        }
        if (capabilities.Contains("officeimo.reader.rtf.pages.explicit", StringComparer.Ordinal)) {
            return OfficeDocumentPageProvenance.ExplicitBreak;
        }

        switch (document.Kind) {
            case ReaderInputKind.Pdf:
                return OfficeDocumentPageProvenance.Native;
            case ReaderInputKind.PowerPoint:
            case ReaderInputKind.Excel:
            case ReaderInputKind.Visio:
            case ReaderInputKind.OneNote:
            case ReaderInputKind.Epub:
                return OfficeDocumentPageProvenance.LogicalContainer;
            default:
                return OfficeDocumentPageProvenance.Unknown;
        }
    }

    private static IReadOnlyList<OfficeDocumentPageLocation> Locate(
        OfficeDocumentReadResult document,
        string blockId,
        OfficeDocumentBlock? fallbackBlock) {
        int totalPageCount = document.GetTotalPageCount();
        OfficeDocumentPageProvenance provenance = document.GetPageProvenance();
        var locations = new List<OfficeDocumentPageLocation>();
        foreach (OfficeDocumentPage page in document.Pages ?? Array.Empty<OfficeDocumentPage>()) {
            OfficeDocumentBlock[] matches = (page.Blocks ?? Array.Empty<OfficeDocumentBlock>())
                .Where(candidate =>
                    !string.IsNullOrEmpty(blockId)
                        ? string.Equals(candidate.Id, blockId, StringComparison.Ordinal)
                        : ReferenceEquals(candidate, fallbackBlock))
                .ToArray();
            if (matches.Length == 0) {
                continue;
            }

            OfficeDocumentRegion[] regions = matches
                .Where(static candidate => candidate.Region != null)
                .Select(static candidate => candidate.Region!)
                .ToArray();
            locations.Add(new OfficeDocumentPageLocation(page, totalPageCount, provenance, regions));
        }

        return locations.AsReadOnly();
    }
}
