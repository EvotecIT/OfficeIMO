using System.Globalization;
using System.Net;

namespace OfficeIMO.Pdf;

/// <summary>Rendered visual and structural comparison report for two PDFs.</summary>
public sealed class PdfVisualComparisonReport {
    internal PdfVisualComparisonReport(IReadOnlyList<PdfVisualPageComparison> pages, IReadOnlyList<string> structuralDifferences) {
        Pages = pages.ToArray();
        StructuralDifferences = structuralDifferences.ToArray();
    }

    /// <summary>Per-page comparisons.</summary>
    public IReadOnlyList<PdfVisualPageComparison> Pages { get; }
    /// <summary>Document/page structural differences.</summary>
    public IReadOnlyList<string> StructuralDifferences { get; }
    /// <summary>True when all compared pages satisfy thresholds and no structural differences remain.</summary>
    public bool IsMatch => StructuralDifferences.Count == 0 && Pages.All(static page => page.IsMatch);

    /// <summary>Builds a self-contained HTML human-review gallery with expected, actual, and highlighted diff images.</summary>
    public string ToHtmlGallery(string? title = null) {
        var html = new StringBuilder("<!doctype html><html><head><meta charset=\"utf-8\"><style>body{font-family:sans-serif}section{margin:1rem 0}.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}img{max-width:100%;border:1px solid #ccc}.fail{color:#b00020}</style></head><body>");
        html.Append("<h1>").Append(WebUtility.HtmlEncode(title ?? "PDF visual comparison")).Append("</h1>");
        foreach (string difference in StructuralDifferences) html.Append("<p class=\"fail\">").Append(WebUtility.HtmlEncode(difference)).Append("</p>");
        foreach (PdfVisualPageComparison page in Pages) {
            html.Append("<section><h2>Page ").Append(page.PageNumber.ToString(CultureInfo.InvariantCulture)).Append(page.IsMatch ? " - match" : " - differs").Append("</h2><p>")
                .Append(page.DifferentPixels.ToString(CultureInfo.InvariantCulture)).Append(" changed pixels; ratio ").Append(page.DifferenceRatio.ToString("0.######", CultureInfo.InvariantCulture)).Append("</p><div class=\"grid\">");
            AppendImage(html, "Expected", page.ExpectedPng);
            AppendImage(html, "Actual", page.ActualPng);
            AppendImage(html, "Diff", page.DiffPng);
            html.Append("</div></section>");
        }

        return html.Append("</body></html>").ToString();
    }

    private static void AppendImage(StringBuilder html, string label, byte[] bytes) {
        html.Append("<figure><figcaption>").Append(label).Append("</figcaption><img alt=\"").Append(label).Append("\" src=\"data:image/png;base64,")
            .Append(Convert.ToBase64String(bytes)).Append("\"></figure>");
    }
}

/// <summary>One rendered page comparison and its human-review artifacts.</summary>
public sealed class PdfVisualPageComparison {
    private readonly byte[] _expectedPng;
    private readonly byte[] _actualPng;
    private readonly byte[] _diffPng;

    internal PdfVisualPageComparison(int pageNumber, bool isMatch, int width, int height, long comparedPixels, long differentPixels, int maximumChannelDifference, double meanChannelDifference, byte[] expectedPng, byte[] actualPng, byte[] diffPng) {
        PageNumber = pageNumber; IsMatch = isMatch; Width = width; Height = height; ComparedPixels = comparedPixels; DifferentPixels = differentPixels;
        MaximumChannelDifference = maximumChannelDifference; MeanChannelDifference = meanChannelDifference;
        _expectedPng = (byte[])expectedPng.Clone(); _actualPng = (byte[])actualPng.Clone(); _diffPng = (byte[])diffPng.Clone();
    }
    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }
    /// <summary>Whether this page satisfies the configured threshold.</summary>
    public bool IsMatch { get; }
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
}
