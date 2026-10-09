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
        string heading = title ?? "PDF visual comparison";
        html.Append("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"><title>");
        html.AppendHtmlEncoded(heading);
        html.Append("</title><style>").Append(GalleryStyles).Append("</style></head><body><main>");
        int differing = Pages.Count(static page => !page.IsMatch);
        html.Append("<header><p class=\"eyebrow\">PDF visual comparison</p><h1>");
        html.AppendHtmlEncoded(heading);
        html.Append("</h1><p class=\"summary\"><span class=\"badge ").Append(IsMatch ? "ok\">Match" : "bad\">Different").Append("</span>")
            .Append(Pages.Count.ToString(CultureInfo.InvariantCulture)).Append(Pages.Count == 1 ? " page compared" : " pages compared")
            .Append(" · ").Append((Pages.Count - differing).ToString(CultureInfo.InvariantCulture)).Append(" matching · ")
            .Append(differing.ToString(CultureInfo.InvariantCulture)).Append(" different</p>");
        if (differing > 0) {
            html.Append("<p class=\"jump\">Go to ");
            int listed = 0;
            foreach (PdfVisualPageComparison page in Pages) {
                if (page.IsMatch) continue;
                if (listed == 20) { html.Append(" …"); break; }
                cancellationToken.ThrowIfCancellationRequested();
                string number = page.PageNumber.ToString(CultureInfo.InvariantCulture);
                html.Append(listed++ == 0 ? string.Empty : " ").Append("<a href=\"#page-").Append(number).Append("\">page ").Append(number).Append("</a>");
            }
            html.Append("</p>");
        }
        html.Append("</header>");
        if (StructuralDifferences.Count > 0) {
            html.Append("<section class=\"alert\"><h2>Structural differences</h2><ul>");
            foreach (string difference in StructuralDifferences) {
                cancellationToken.ThrowIfCancellationRequested();
                html.Append("<li>");
                html.AppendHtmlEncoded(difference);
                html.Append("</li>");
            }
            html.Append("</ul></section>");
        }
        foreach (PdfVisualPageComparison page in Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            html.Append("<section id=\"page-").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture)).Append("\" class=\"page").Append(page.IsMatch ? "\">" : " differs\">").Append("<div class=\"page-head\"><h2>Page ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture));
            if (page.ActualPageNumber != page.PageNumber) html.Append(" vs. ").Append(page.ActualPageNumber.ToString(CultureInfo.InvariantCulture));
            html.Append("</h2><span class=\"badge ").Append(page.IsMatch ? "ok\">Match" : "bad\">Differs").Append("</span><p>")
                .Append(page.DifferentPixels.ToString("N0", CultureInfo.InvariantCulture)).Append(" changed pixels · ")
                .Append((page.DifferenceRatio * 100D).ToString("0.###", CultureInfo.InvariantCulture)).Append("% of the page</p></div>");
            if (PdfRenderCapabilities.HasIncompleteVisualProjection(page.ExpectedCapabilityDiagnostics) ||
                PdfRenderCapabilities.HasIncompleteVisualProjection(page.ActualCapabilityDiagnostics)) {
                html.Append("<p class=\"note\">Managed rendering is incomplete; these pixel images alone cannot prove a match.</p>");
            }
            html.Append("<div class=\"grid\">");
            AppendImage(html, "Expected", page.ExpectedPng, cancellationToken);
            AppendImage(html, "Actual", page.ActualPng, cancellationToken);
            AppendImage(html, "Changes", page.DiffPng, cancellationToken);
            html.Append("</div></section>");
        }

        return html.Append("</main></body></html>").ToString();
    }

    // Self-contained styles: readable on phones, light or dark with the viewer's preference, no external assets.
    private const string GalleryStyles =
        ":root{color-scheme:light dark;--bg:#f5f7fb;--card:#fff;--line:#dde3ec;--text:#0d1526;--muted:#5b6880;--ok:#047857;--bad:#be123c;--warn:#92400e}" +
        "@media (prefers-color-scheme:dark){:root{--bg:#0b1020;--card:#131a2e;--line:#26304a;--text:#e7edf6;--muted:#97a3ba;--ok:#34d399;--bad:#fb7185;--warn:#fbbf24}}" +
        "*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--text);font:15px/1.55 system-ui,-apple-system,'Segoe UI',Roboto,sans-serif}" +
        "main{max-width:1180px;margin:0 auto;padding:28px 20px 48px}header{margin-bottom:20px}" +
        ".eyebrow{margin:0 0 6px;font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}" +
        "h1{margin:0 0 10px;font-size:clamp(22px,4vw,32px);line-height:1.2}h2{margin:0;font-size:17px}" +
        ".summary{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:0;color:var(--muted)}" +
        ".badge{display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px;font-weight:700;color:var(--ok);background:color-mix(in srgb,var(--ok) 14%,transparent)}" +
        ".badge.bad{color:var(--bad);background:color-mix(in srgb,var(--bad) 14%,transparent)}" +
        ".alert,.page{margin:0 0 14px;padding:16px 18px;border:1px solid var(--line);border-radius:14px;background:var(--card)}" +
        ".alert{border-color:color-mix(in srgb,var(--bad) 45%,var(--line))}.alert ul{margin:8px 0 0;padding-left:20px}" +
        ".page.differs{border-color:color-mix(in srgb,var(--bad) 40%,var(--line))}" +
        ".page-head{display:flex;flex-wrap:wrap;align-items:center;gap:6px 10px;margin-bottom:12px}.page-head p{flex-basis:100%;margin:0;font-size:13px;color:var(--muted)}" +
        ".note{margin:0 0 12px;font-size:13px;color:var(--warn)}" +
        ".jump{margin:10px 0 0;font-size:14px;color:var(--muted)}.jump a{margin-right:4px;padding:2px 9px;border-radius:999px;color:var(--bad);text-decoration:none;font-weight:600;background:color-mix(in srgb,var(--bad) 12%,transparent)}" +
        ".grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}" +
        "figure{margin:0}figcaption{margin-bottom:6px;font-size:12px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:var(--muted)}" +
        "img{display:block;width:100%;height:auto;border:1px solid var(--line);border-radius:8px;background:#fff}" +
        "@media (max-width:640px){main{padding:18px 14px 32px}.grid{grid-template-columns:1fr 1fr}.grid figure:last-child{grid-column:1/-1}}";

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
            int firstChangedX = bounds.X;
            int right = bounds.X + bounds.Width;
            while (firstChangedX < right && !_changedPixels[checked(y * Width + firstChangedX)]) firstChangedX++;
            if (firstChangedX == right) continue;
            intervals.Clear();
            for (int index = 0; index < classified.Count; index++) {
                PdfPixelRegion region = classified[index];
                if (region.Y <= y && y < region.Y + region.Height) intervals.Add(region);
            }
            intervals.Sort(static (left, right) => left.X.CompareTo(right.X));
            int intervalIndex = 0;
            int coveredRight = 0;
            for (int x = firstChangedX; x < right; x++) {
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

    internal bool HasChangedPixelsIn(PdfPixelRegion region, System.Threading.CancellationToken cancellationToken) {
        int right = Math.Min(Width, region.X + region.Width);
        int bottom = Math.Min(Height, region.Y + region.Height);
        for (int y = Math.Max(0, region.Y); y < bottom; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = Math.Max(0, region.X); x < right; x++) {
                if (_changedPixels[checked(y * Width + x)]) return true;
            }
        }
        return false;
    }
}
