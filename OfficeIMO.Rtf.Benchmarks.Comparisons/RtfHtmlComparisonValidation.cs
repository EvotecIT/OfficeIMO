using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
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
            RequireMarker(office, "Status");
            RequireMarker(rtfPipe, "Status");
            RequireMarker(office, "RTF to HTML");
            RequireMarker(rtfPipe, "RTF to HTML");
            RequireMarker(office, "Verified");
            RequireMarker(rtfPipe, "Verified");
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

        RequireEquivalentSemanticContent(scale, office, rtfPipe);

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
        string text = ExtractSemanticText(document);
        if (text.Length == 0) {
            throw new InvalidOperationException($"{implementation} produced HTML without visible text for '{scale}'.");
        }

        string[] semanticTokens = GetSemanticTokens(text);
        string[] tableCells = document.QuerySelectorAll("th,td")
            .Select(cell => NormalizeWhitespace(cell.TextContent))
            .ToArray();

        return new RtfHtmlOutputEvidence(
            implementation,
            Encoding.UTF8.GetByteCount(html),
            RecordMarkerRegex().Matches(text).Count,
            document.QuerySelectorAll("table").Length,
            tableCells.Length,
            document.QuerySelectorAll("img").Length,
            semanticTokens.Length,
            Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join('\u001F', semanticTokens)))),
            tableCells,
            text);
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
        for (int rowIndex = 1; rowIndex <= 3; rowIndex++) {
            for (int columnIndex = 1; columnIndex <= 3; columnIndex++) {
                RequireMarker(evidence, $"R{rowIndex} C{columnIndex}");
            }
        }
    }

    private static void RequireEquivalentSemanticContent(
        string scale,
        RtfHtmlOutputEvidence office,
        RtfHtmlOutputEvidence rtfPipe) {
        string[] officeTokens = GetSemanticTokens(office.Text);
        string[] rtfPipeTokens = GetSemanticTokens(rtfPipe.Text);
        if (!officeTokens.SequenceEqual(rtfPipeTokens, StringComparer.Ordinal)) {
            throw new InvalidOperationException(
                $"RTF-to-HTML semantic text differs for '{scale}': " +
                $"OfficeIMO {officeTokens.Length} tokens/{office.SemanticSha256}; " +
                $"RtfPipe {rtfPipeTokens.Length} tokens/{rtfPipe.SemanticSha256}.");
        }

        if (office.TableCells.Count != rtfPipe.TableCells.Count) {
            throw new InvalidOperationException($"RTF-to-HTML table cell counts differ for '{scale}'.");
        }

        for (int index = 0; index < office.TableCells.Count; index++) {
            string[] officeCellTokens = GetSemanticTokens(office.TableCells[index]);
            string[] rtfPipeCellTokens = GetSemanticTokens(rtfPipe.TableCells[index]);
            if (!officeCellTokens.SequenceEqual(rtfPipeCellTokens, StringComparer.Ordinal)) {
                throw new InvalidOperationException(
                    $"RTF-to-HTML table cell {index + 1} differs for '{scale}'.");
            }
        }
    }

    private static void RequireMarker(RtfHtmlOutputEvidence evidence, string marker) {
        if (!evidence.Text.Contains(marker, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{evidence.Implementation} did not preserve required marker '{marker}'.");
        }
    }

    private static string NormalizeWhitespace(string value) =>
        WhitespaceRegex().Replace(value, " ").Trim();

    private static string ExtractSemanticText(IHtmlDocument document) {
        var builder = new StringBuilder();
        INode? root = document.Body ?? document.DocumentElement;
        if (root != null) AppendSemanticText(root, builder);
        return NormalizeWhitespace(builder.ToString());
    }

    private static void AppendSemanticText(INode node, StringBuilder builder) {
        if (node is IText text) {
            builder.Append(text.Data);
            return;
        }

        bool boundary = node is IElement element && IsSemanticBlockBoundary(element.TagName);
        if (boundary) builder.Append(' ');
        foreach (INode child in node.ChildNodes) AppendSemanticText(child, builder);
        if (boundary) builder.Append(' ');
    }

    private static bool IsSemanticBlockBoundary(string tagName) => tagName switch {
        "ADDRESS" or "ARTICLE" or "ASIDE" or "BLOCKQUOTE" or "BR" or "CAPTION" or
        "DD" or "DIV" or "DL" or "DT" or "FIGCAPTION" or "FIGURE" or "FOOTER" or
        "H1" or "H2" or "H3" or "H4" or "H5" or "H6" or "HEADER" or "HR" or
        "LI" or "MAIN" or "NAV" or "OL" or "P" or "PRE" or "SECTION" or "TABLE" or
        "TBODY" or "TD" or "TFOOT" or "TH" or "THEAD" or "TR" or "UL" => true,
        _ => false
    };

    private static string[] GetSemanticTokens(string value) =>
        SemanticTokenRegex()
            .Matches(value ?? string.Empty)
            .Select(match => match.Value)
            .ToArray();

    [GeneratedRegex(@"\bRecord\s+\d+:", RegexOptions.CultureInvariant)]
    private static partial Regex RecordMarkerRegex();

    [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant)]
    private static partial Regex WhitespaceRegex();

    [GeneratedRegex(@"[\p{L}\p{M}\p{N}]+|[^\s\p{L}\p{M}\p{N}]", RegexOptions.CultureInvariant)]
    private static partial Regex SemanticTokenRegex();
}
