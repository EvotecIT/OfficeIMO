using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Rtf.Benchmarks.Comparisons;

internal static partial class RtfHtmlComparisonValidation {
    internal static IReadOnlyList<RtfHtmlComparisonReport> ValidateAll() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return RtfHtmlComparisonCorpus.Scales
            .Select(scale => Validate(scale, RtfHtmlComparisonCorpus.Get(scale).Rtf))
            .ToArray();
    }

    internal static RtfHtmlComparisonReport Validate(string scale, string rtf) {
        string officeHtml = RtfDocument.Read(rtf).Document.ToHtml();
        string rtfPipeHtml = global::RtfPipe.Rtf.ToHtml(rtf);
        RtfHtmlOutputEvidence office = Inspect("OfficeIMO", scale, officeHtml);
        RtfHtmlOutputEvidence rtfPipe = Inspect("RtfPipe", scale, rtfPipeHtml);

        RtfHtmlComparisonFixture fixture = RtfHtmlComparisonCorpus.Get(scale);
        if (string.Equals(scale, "Producer", StringComparison.OrdinalIgnoreCase)) {
            RequireMarker(office, "Commercial library RTF fixture");
            RequireMarker(rtfPipe, "Commercial library RTF fixture");
            RequireMarker(office, "Route");
            RequireMarker(rtfPipe, "Route");
        } else {
            RequireGeneratedCorpus(office, fixture.ParagraphCount);
            RequireGeneratedCorpus(rtfPipe, fixture.ParagraphCount);
        }

        if (office.TableCount != rtfPipe.TableCount || office.CellCount != rtfPipe.CellCount) {
            throw new InvalidOperationException(
                $"RTF-to-HTML table structure differs for '{scale}': " +
                $"OfficeIMO {office.TableCount} tables/{office.CellCount} cells; " +
                $"RtfPipe {rtfPipe.TableCount} tables/{rtfPipe.CellCount} cells.");
        }

        return new RtfHtmlComparisonReport(
            scale,
            Encoding.UTF8.GetByteCount(rtf),
            office,
            rtfPipe);
    }

    private static RtfHtmlOutputEvidence Inspect(string implementation, string scale, string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            throw new InvalidOperationException($"{implementation} produced empty HTML for '{scale}'.");
        }

        IHtmlDocument document = HtmlConversionDocument.Parse(html).CreateDocumentForConversion();
        string text = NormalizeWhitespace(document.Body?.TextContent ?? document.DocumentElement?.TextContent ?? string.Empty);
        if (text.Length == 0) {
            throw new InvalidOperationException($"{implementation} produced HTML without visible text for '{scale}'.");
        }

        return new RtfHtmlOutputEvidence(
            implementation,
            Encoding.UTF8.GetByteCount(html),
            text,
            RecordMarkerRegex().Matches(text).Count,
            document.QuerySelectorAll("table").Length,
            document.QuerySelectorAll("th,td").Length,
            document.QuerySelectorAll("img").Length);
    }

    private static void RequireGeneratedCorpus(RtfHtmlOutputEvidence evidence, int expectedRecords) {
        if (evidence.RecordCount != expectedRecords) {
            throw new InvalidOperationException(
                $"{evidence.Implementation} preserved {evidence.RecordCount} of {expectedRecords} records.");
        }

        RequireMarker(evidence, "Record 1:");
        RequireMarker(evidence, $"Record {expectedRecords}:");
        RequireMarker(evidence, "zażółć gęślą jaźń");
        if (expectedRecords > 1) RequireMarker(evidence, "Καλημέρα Привет");
        RequireMarker(evidence, "R1 C1");
    }

    private static void RequireMarker(RtfHtmlOutputEvidence evidence, string marker) {
        if (!evidence.Text.Contains(marker, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{evidence.Implementation} did not preserve required marker '{marker}'.");
        }
    }

    private static string NormalizeWhitespace(string value) =>
        WhitespaceRegex().Replace(value, " ").Trim();

    [GeneratedRegex(@"\bRecord\s+\d+:", RegexOptions.CultureInvariant)]
    private static partial Regex RecordMarkerRegex();

    [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant)]
    private static partial Regex WhitespaceRegex();
}
