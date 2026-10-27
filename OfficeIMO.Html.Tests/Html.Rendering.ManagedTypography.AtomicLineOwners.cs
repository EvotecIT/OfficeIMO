using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Text.RegularExpressions;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("inline-block", false)]
    [InlineData("inline-block", true)]
    [InlineData("inline-table", false)]
    [InlineData("inline-table", true)]
    [InlineData("inline-flex", false)]
    [InlineData("inline-flex", true)]
    [InlineData("inline-grid", false)]
    [InlineData("inline-grid", true)]
    [InlineData("ruby", false)]
    [InlineData("ruby", true)]
    public void HtmlManagedTypography_LineLogicalOwnerRetainsAtomicChildren(string kind, bool explicitProvider) {
        byte[] data = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf"));
        string atomic = kind switch {
            "inline-table" => "<span id='atomic' style='display:inline-table;border-spacing:0'><span style='display:table-row'><span style='display:table-cell;padding:0;border:0'>INNER</span></span></span>",
            "ruby" => "<ruby id='atomic'><rb>INNER</rb><rt>NOTE</rt></ruby>",
            _ => "<span id='atomic' style='display:" + kind + "'>INNER</span>"
        };
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@page{size:420px 160px;margin:0}html,body,p{margin:0;padding:0}"
            + "p{font:16px/28px AtomicTypography}table{border-spacing:0}td{padding:0;border:0}"
            + "ruby{ruby-align:start}</style><p>Cafe\u0301 " + atomic + " tail</p>");
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            FontFamily = new PdfCore.PdfEmbeddedFontFamily("AtomicTypography", data)
        };
        options.Fonts.Add("AtomicTypography", data);
        if (explicitProvider) options.TextShapingProvider = OfficeManagedTextShapingProvider.Instance;
        string[] expected = kind == "ruby" ? new[] { "Cafe\u0301", "INNER", "NOTE", "tail" }
            : new[] { "Cafe\u0301", "INNER", "tail" };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options);

        rendered.RequireNoLoss();
        Assert.Single(rendered.Pages);
        AssertAtomicLineLogicalTokens(rendered.Text, expected);
        PdfCore.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(options);
        pdf.RequireNoLoss();
        byte[] bytes = pdf.ToBytes();
        PdfCore.PdfReadDocument saved = PdfCore.PdfReadDocument.Open(bytes);
        Assert.Single(saved.Pages);
        // Spatial extraction orders independently placed boxes by their baselines.
        // Tagged reading order retains the authored sequence, including ruby.
        AssertAtomicLineLogicalTokens(saved.ExtractText(), expected, sourceOrder: false);
        var tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(bytes).TaggedContent);
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        var root = Assert.Single(tagged.StructureElements, item => item.StructureType == "Document");
        AssertAtomicLineLogicalTokens(string.Concat(SourceOrderOwnedText(bytes, root, elements)), expected);
        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options);
        OfficeImageExportResult replay = Assert.Single(PdfCore.PdfImageExportExtensions.ExportImages(
            pdf, OfficeImageExportFormat.Png, new PdfCore.PdfImageExportOptions { Scale = 96D / 72D }));
        Assert.Equal(replay.Bytes, png.Bytes);
        Assert.DoesNotContain(png.Diagnostics, item => item.LossKind != OfficeConversionLossKind.None);
    }

    private static void AssertAtomicLineLogicalTokens(string logicalText, string[] expected, bool sourceOrder = true) {
        // Block/annotation separators may be explicit. Their spelling does not
        // permit an atomic child's content to disappear, repeat, or move order.
        const string pattern = "Cafe\u0301|INNER|NOTE|tail";
        IEnumerable<string> actual = Regex.Matches(logicalText, pattern).Cast<Match>().Select(match => match.Value);
        Assert.Equal(sourceOrder ? expected : expected.OrderBy(value => value, StringComparer.Ordinal),
            sourceOrder ? actual : actual.OrderBy(value => value, StringComparer.Ordinal));
        string remaining = Regex.Replace(logicalText, pattern, string.Empty);
        Assert.True(remaining.All(char.IsWhiteSpace), "Unexpected logical text: " + logicalText);
    }
}
