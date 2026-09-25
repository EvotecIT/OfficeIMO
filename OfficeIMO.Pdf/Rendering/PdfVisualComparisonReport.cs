using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Rendered visual and structural comparison report for two PDFs.</summary>
public sealed class PdfVisualComparisonReport {
    internal PdfVisualComparisonReport(IReadOnlyList<PdfVisualPageComparison> pages, IReadOnlyList<string> structuralDifferences, int expectedPageCount, int actualPageCount) {
        Pages = pages.ToArray();
        StructuralDifferences = structuralDifferences.ToArray();
        ExpectedPageCount = expectedPageCount;
        ActualPageCount = actualPageCount;
    }

    /// <summary>Total pages in the expected document, including pages outside a selected comparison.</summary>
    public int ExpectedPageCount { get; }
    /// <summary>Total pages in the actual document, including pages outside a selected comparison.</summary>
    public int ActualPageCount { get; }

    /// <summary>Per-page comparisons.</summary>
    public IReadOnlyList<PdfVisualPageComparison> Pages { get; }
    /// <summary>Document/page structural differences.</summary>
    public IReadOnlyList<string> StructuralDifferences { get; }
    /// <summary>True when all compared pages have complete managed renderings that satisfy thresholds and no structural differences remain.</summary>
    public bool IsMatch => StructuralDifferences.Count == 0 && Pages.All(static page => page.IsMatch);

    /// <summary>Builds a self-contained HTML human-review gallery with expected, actual, and highlighted diff images.</summary>
    public string ToHtmlGallery(string? title = null) =>
        ToHtmlGallery(title, long.MaxValue, CancellationToken.None);

    /// <summary>Builds a self-contained HTML gallery without allowing its UTF-8 representation to exceed the supplied byte limit.</summary>
    public string ToHtmlGallery(
        string? title,
        long maximumOutputBytes,
        CancellationToken cancellationToken = default) {
        var html = new BoundedUtf8HtmlBuilder(maximumOutputBytes);
        html.Append("<!doctype html><html><head><meta charset=\"utf-8\"><style>body{font-family:sans-serif}section{margin:1rem 0}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}img{max-width:100%;border:1px solid #ccc}.fail{color:#b00020}</style></head><body>");
        html.Append("<h1>");
        html.AppendHtmlEncoded(title ?? "PDF visual comparison");
        html.Append("</h1>");
        foreach (string difference in StructuralDifferences) {
            cancellationToken.ThrowIfCancellationRequested();
            html.Append("<p class=\"fail\">");
            html.AppendHtmlEncoded(difference);
            html.Append("</p>");
        }
        foreach (PdfVisualPageComparison page in Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            html.Append("<section><h2>Page ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture));
            if (page.ActualPageNumber != page.PageNumber) html.Append(" vs. ").Append(page.ActualPageNumber.ToString(CultureInfo.InvariantCulture));
            html.Append(page.IsMatch ? " - match" : " - differs").Append("</h2><p>")
                .Append(page.DifferentPixels.ToString(CultureInfo.InvariantCulture)).Append(" changed pixels; ratio ").Append(page.DifferenceRatio.ToString("0.######", CultureInfo.InvariantCulture)).Append("</p>");
            if (PdfRenderCapabilities.HasIncompleteVisualProjection(page.ExpectedCapabilityDiagnostics) ||
                PdfRenderCapabilities.HasIncompleteVisualProjection(page.ActualCapabilityDiagnostics)) {
                html.Append("<p class=\"fail\">Managed rendering is incomplete; these pixel images alone cannot prove a match.</p>");
            }
            html.Append("<div class=\"grid\">");
            AppendImage(html, "Expected", page.ExpectedPng, cancellationToken);
            AppendImage(html, "Actual", page.ActualPng, cancellationToken);
            AppendImage(html, "Diff", page.DiffPng, cancellationToken);
            html.Append("</div></section>");
        }

        return html.Append("</body></html>").ToString();
    }

    private static void AppendImage(
        BoundedUtf8HtmlBuilder html,
        string label,
        byte[] bytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        html.Append("<figure><figcaption>").Append(label).Append("</figcaption><img alt=\"").Append(label).Append("\" src=\"data:image/png;base64,")
            .AppendBase64(bytes).Append("\"></figure>");
    }

    private sealed class BoundedUtf8HtmlBuilder {
        private readonly StringBuilder _builder;
        private readonly long _maximumOutputBytes;
        private long _outputBytes;

        internal BoundedUtf8HtmlBuilder(long maximumOutputBytes) {
#pragma warning disable CA1512 // ThrowIfLessThanOrEqual is unavailable on netstandard2.0 and net472.
            if (maximumOutputBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
#pragma warning restore CA1512
            _maximumOutputBytes = maximumOutputBytes;
            _builder = new StringBuilder((int)Math.Min(maximumOutputBytes, 4096L));
        }

        internal BoundedUtf8HtmlBuilder Append(string value) {
            if (string.IsNullOrEmpty(value)) return this;
            Charge(Encoding.UTF8.GetByteCount(value));
            _builder.Append(value);
            return this;
        }

        internal BoundedUtf8HtmlBuilder AppendHtmlEncoded(string value) {
            if (string.IsNullOrEmpty(value)) return this;
            int unencodedStart = 0;
            for (int index = 0; index < value.Length; index++) {
                string? entity = value[index] switch {
                    '&' => "&amp;",
                    '<' => "&lt;",
                    '>' => "&gt;",
                    '"' => "&quot;",
                    '\'' => "&#39;",
                    _ => null
                };
                if (entity == null) continue;
                Append(value, unencodedStart, index - unencodedStart);
                Append(entity);
                unencodedStart = index + 1;
            }
            Append(value, unencodedStart, value.Length - unencodedStart);
            return this;
        }

        internal BoundedUtf8HtmlBuilder AppendBase64(byte[] bytes) {
            long encodedLength = checked(4L * ((bytes.LongLength + 2L) / 3L));
            Charge(encodedLength);
            _builder.Append(Convert.ToBase64String(bytes));
            return this;
        }

        public override string ToString() => _builder.ToString();

        private void Append(string value, int startIndex, int count) {
            if (count == 0) return;
            var buffer = new char[Math.Min(count, 1024)];
            int sourceIndex = startIndex;
            int remaining = count;
            while (remaining > 0) {
                int chunkLength = Math.Min(buffer.Length, remaining);
                if (chunkLength < remaining &&
                    char.IsHighSurrogate(value[sourceIndex + chunkLength - 1]) &&
                    char.IsLowSurrogate(value[sourceIndex + chunkLength])) {
                    chunkLength--;
                }
                value.CopyTo(sourceIndex, buffer, 0, chunkLength);
                Charge(Encoding.UTF8.GetByteCount(buffer, 0, chunkLength));
                sourceIndex += chunkLength;
                remaining -= chunkLength;
            }
            _builder.Append(value, startIndex, count);
        }

        private void Charge(long byteCount) {
            if (byteCount > _maximumOutputBytes - _outputBytes) ThrowLimitExceeded();
            _outputBytes += byteCount;
        }

        private void ThrowLimitExceeded() => throw new InvalidOperationException(
            $"Generated comparison gallery exceeded the configured {_maximumOutputBytes:N0}-byte output limit while it was being rendered.");
    }
}

/// <summary>One rendered page comparison and its human-review artifacts.</summary>
public sealed class PdfVisualPageComparison {
    private readonly System.Collections.BitArray _changedPixels;
    private readonly byte[] _expectedPng;
    private readonly byte[] _actualPng;
    private readonly byte[] _diffPng;

    internal PdfVisualPageComparison(int pageNumber, int actualPageNumber, bool isMatch, int width, int height, long comparedPixels, long differentPixels, int maximumChannelDifference, double meanChannelDifference, byte[] expectedPng, byte[] actualPng, byte[] diffPng, bool hasSizeDifference, PdfPixelRegion? changedBounds, System.Collections.BitArray changedPixels, IReadOnlyList<PdfRenderCapabilityDiagnostic> expectedDiagnostics, IReadOnlyList<PdfRenderCapabilityDiagnostic> actualDiagnostics) {
        PageNumber = pageNumber; ActualPageNumber = actualPageNumber; IsMatch = isMatch; Width = width; Height = height; ComparedPixels = comparedPixels; DifferentPixels = differentPixels;
        MaximumChannelDifference = maximumChannelDifference; MeanChannelDifference = meanChannelDifference;
        HasSizeDifference = hasSizeDifference; ChangedBounds = changedBounds;
        ExpectedCapabilityDiagnostics = Array.AsReadOnly(expectedDiagnostics.ToArray());
        ActualCapabilityDiagnostics = Array.AsReadOnly(actualDiagnostics.ToArray());
        _expectedPng = (byte[])expectedPng.Clone(); _actualPng = (byte[])actualPng.Clone(); _diffPng = (byte[])diffPng.Clone();
        _changedPixels = (System.Collections.BitArray)changedPixels.Clone();
    }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>One-based page in the actual document.</summary>
    public int ActualPageNumber { get; }
    /// <summary>Whether this page has complete managed renderings that satisfy the configured threshold.</summary>
    public bool IsMatch { get; }
    /// <summary>Known simplifications or omissions in the expected page's managed rendering.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> ExpectedCapabilityDiagnostics { get; }
    /// <summary>Known simplifications or omissions in the actual page's managed rendering.</summary>
    public IReadOnlyList<PdfRenderCapabilityDiagnostic> ActualCapabilityDiagnostics { get; }
    /// <summary>Whether the rendered source dimensions differ, independently of pixel tolerances.</summary>
    public bool HasSizeDifference { get; }
    /// <summary>Smallest comparison-canvas rectangle containing pixels above channel tolerance, or null when none differ.</summary>
    public PdfPixelRegion? ChangedBounds { get; }
    /// <summary>Comparison canvas width.</summary>
    public int Width { get; }
    /// <summary>Comparison canvas height.</summary>
    public int Height { get; }
    /// <summary>Pixels compared after exclusions.</summary>
    public long ComparedPixels { get; }
    /// <summary>Pixels exceeding channel tolerance.</summary>
    public long DifferentPixels { get; }
    /// <summary>Maximum observed channel difference.</summary>
    public int MaximumChannelDifference { get; }
    /// <summary>Mean absolute channel difference.</summary>
    public double MeanChannelDifference { get; }
    /// <summary>Changed-pixel ratio.</summary>
    public double DifferenceRatio => ComparedPixels == 0 ? 0D : DifferentPixels / (double)ComparedPixels;
    /// <summary>Expected page PNG.</summary>
    public byte[] ExpectedPng => (byte[])_expectedPng.Clone();
    /// <summary>Actual page PNG.</summary>
    public byte[] ActualPng => (byte[])_actualPng.Clone();
    /// <summary>Highlighted diff PNG.</summary>
    public byte[] DiffPng => (byte[])_diffPng.Clone();
    internal long OutputByteLength => checked(_expectedPng.LongLength + _actualPng.LongLength + _diffPng.LongLength);
    internal bool HasChangedPixelsOutside(IReadOnlyList<PdfPixelRegion> classified, System.Threading.CancellationToken cancellationToken) {
        if (ChangedBounds is not PdfPixelRegion bounds) return false;
        var intervals = new List<PdfPixelRegion>(classified.Count);
        for (int y = bounds.Y; y < bounds.Y + bounds.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            intervals.Clear();
            for (int index = 0; index < classified.Count; index++) {
                PdfPixelRegion region = classified[index];
                if (region.Y <= y && y < region.Y + region.Height) intervals.Add(region);
            }
            intervals.Sort(static (left, right) => left.X.CompareTo(right.X));
            int intervalIndex = 0;
            int coveredRight = 0;
            for (int x = bounds.X; x < bounds.X + bounds.Width; x++) {
                if (!_changedPixels[checked(y * Width + x)]) continue;
                while (intervalIndex < intervals.Count && intervals[intervalIndex].X <= x) {
                    PdfPixelRegion region = intervals[intervalIndex++];
                    coveredRight = Math.Max(coveredRight, region.X + region.Width);
                }
                if (x >= coveredRight) return true;
            }
        }
        return false;
    }
}
