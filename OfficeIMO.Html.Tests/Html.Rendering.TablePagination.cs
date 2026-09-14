using System.IO.Compression;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTables_CssRowGroupsControlPagedHeaderAndFooterRepetition() {
        string bodyRows = string.Concat(Enumerable.Range(0, 14).Select(index =>
            "<tr><td>Body" + index.ToString("D2") + "</td><td>Value</td></tr>"));
        string html = "<style>@page{size:180px 80px;margin:0}table{width:160px;margin:0;border-collapse:collapse}"
            + "th,td{font-size:8px;line-height:10px;padding:2px;border:1px solid #456}"
            + "thead{display:table-row-group}.repeat{display:table-header-group}.ending{display:table-footer-group}</style>"
            + "<table id='report'><thead><tr><th>NativeOnce</th><th>Value</th></tr></thead>"
            + "<tbody class='repeat'><tr><th id='css-header-name'>CssHeader</th><th id='css-header-value'>Value</th></tr></tbody>"
            + "<tbody>" + bodyRows + "</tbody>"
            + "<tbody class='ending'><tr><td>CssFooter</td><td>End</td></tr></tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        IReadOnlyList<HtmlRenderPage> bodyPages = rendered.Pages
            .Where(page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text.StartsWith("Body", StringComparison.Ordinal)))
            .ToList();

        Assert.True(bodyPages.Count >= 3);
        Assert.All(bodyPages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "CssHeader"));
        Assert.All(bodyPages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "CssFooter"));
        Assert.Equal(1, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count(text => text.Text == "NativeOnce"));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed
            || diagnostic.Code == HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed);

        HtmlRenderSemanticGroup[] repeatedHeaderCells = rendered.Pages
            .SelectMany(page => EnumerateTablePaginationScene(page.Scene))
            .OfType<HtmlRenderSemanticGroup>()
            .Where(group => group.Role == HtmlRenderSemanticGroupRole.TableHeaderCell && group.Source == "th#css-header-name")
            .ToArray();
        Assert.True(repeatedHeaderCells.Length > 2);
        Assert.Single(repeatedHeaderCells.Select(group => group.StructureElementKey).Where(key => key != null).Distinct());

        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(
            PdfCore.PdfInspector.Inspect(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
                PageSize = new OfficePageSize(180D / HtmlRenderOptions.CssPixelsPerInch, 80D / HtmlRenderOptions.CssPixelsPerInch),
                Margins = HtmlRenderMargins.All(0D),
                HonorCssPageRules = true
            })).TaggedContent);
        Assert.Equal(4, tagged.StructureElements.Count(element => element.StructureType == "TH"));
    }

    [Fact]
    public void HtmlTables_RowSpanAndBreakAvoidBoundariesMoveCohesiveGroups() {
        const string html = "<style>@page{size:180px 50px;margin:0}table{margin:0;border-collapse:collapse}"
            + "td{font-size:8px;line-height:14px;padding:2px;border:1px solid #456}tbody.keep{break-inside:avoid}</style>"
            + "<table><tbody><tr><td>Lead</td><td>LeadValue</td></tr></tbody>"
            + "<tbody class='keep'><tr><td id='span' rowspan='2'>SpanGroup</td><td>GroupA</td></tr>"
            + "<tr><td>GroupB</td></tr></tbody><tbody><tr><td>Tail</td><td>TailValue</td></tr></tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderPage groupPage = Assert.Single(rendered.Pages,
            page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "GroupA"));

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "GroupA");
        Assert.Contains(groupPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "GroupB");
        Assert.Contains(groupPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "SpanGroup");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlTables_SingleRowAvoidedGroupMovesInsteadOfSplittingInsideTheRow() {
        const string html = "<style>@page{size:180px 58px;margin:0}table{width:170px;margin:0;border-collapse:collapse}"
            + "td{font-size:8px;line-height:12px;padding:2px;border:1px solid #456}tbody.keep{break-inside:avoid}</style>"
            + "<table><tbody><tr><td>Lead</td><td>LeadValue</td></tr></tbody>"
            + "<tbody class='keep'><tr><td>KeepA1<br>KeepA2<br>KeepA3</td><td>KeepB1<br>KeepB2<br>KeepB3</td></tr></tbody>"
            + "</table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderPage groupPage = Assert.Single(rendered.Pages,
            page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "KeepA1"));

        Assert.NotSame(rendered.Pages[0], groupPage);
        foreach (string marker in new[] { "KeepA1", "KeepA2", "KeepA3", "KeepB1", "KeepB2", "KeepB3" }) {
            Assert.Contains(groupPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == marker);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
    }

    [Fact]
    public void HtmlTables_OversizedMultiCellRowUsesAlignedLineBreaksAcrossOutputs() {
        const string html = "<style>@page{size:180px 58px;margin:0}table{width:170px;margin:0;border-collapse:collapse}"
            + "td{font-size:8px;line-height:14px;padding:2px;border:1px solid #456}</style>"
            + "<table id='matrix'><tr><td>A1<br>A2<br>A3<br>A4<br>A5<br>A6</td>"
            + "<td>B1<br>B2<br>B3<br>B4<br>B5<br>B6</td></tr></table><p style='margin:0'>AfterTable</p>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        string text = rendered.Text;

        Assert.True(rendered.Pages.Count >= 2);
        foreach (string marker in new[] { "A1", "A6", "B1", "B6", "AfterTable" }) {
            Assert.Equal(1, text.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);

        HtmlRenderResult svgResult = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg, options));
        HtmlRenderArchiveResult svgArchive = svgResult.ExportArchive();
        HtmlRenderResult pngResult = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        HtmlRenderArchiveResult pngArchive = pngResult.ExportArchive();
        using var svgZip = new ZipArchive(new MemoryStream(svgArchive.Bytes), ZipArchiveMode.Read);
        using var pngZip = new ZipArchive(new MemoryStream(pngArchive.Bytes), ZipArchiveMode.Read);

        Assert.Equal(rendered.Pages.Count, svgArchive.Manifest.Pages.Count);
        Assert.Equal(rendered.Pages.Count, pngArchive.Manifest.Pages.Count);
        Assert.Equal(rendered.Pages.Count, svgZip.Entries.Count(entry => entry.FullName.EndsWith(".svg", StringComparison.Ordinal)));
        Assert.Equal(rendered.Pages.Count, pngZip.Entries.Count(entry => entry.FullName.EndsWith(".png", StringComparison.Ordinal)));
        string pdfText = PdfCore.PdfReadDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options))).ExtractText();
        Assert.Contains("A1", pdfText, StringComparison.Ordinal);
        Assert.Contains("A6", pdfText, StringComparison.Ordinal);
        Assert.Contains("B1", pdfText, StringComparison.Ordinal);
        Assert.Contains("B6", pdfText, StringComparison.Ordinal);
        Assert.Contains("AfterTable", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTables_ForcedRowBreakAndOversizedAvoidFallbackAreDeterministic() {
        const string forcedHtml = "<style>@page{size:160px 90px;margin:0}table{margin:0}td,th{font-size:8px;line-height:12px;padding:2px}</style>"
            + "<table><thead><tr><th>Header</th></tr></thead><tbody><tr style='break-after:page'><td>First</td></tr>"
            + "<tr><td>Second</td></tr></tbody></table>";
        HtmlRenderDocument forced = HtmlRenderTestDriver.Render(forcedHtml, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, forced.Pages.Count);
        Assert.Contains(forced.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "First");
        Assert.DoesNotContain(forced.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Second");
        Assert.Contains(forced.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Second");
        Assert.All(forced.Pages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header"));

        const string avoidedHtml = "<style>@page{size:160px 45px;margin:0}table{margin:0}td{font-size:8px;line-height:14px;padding:2px}</style>"
            + "<table id='oversized'><tr style='break-inside:avoid'><td>A1<br>A2<br>A3<br>A4</td><td>B1<br>B2<br>B3<br>B4</td></tr></table>";
        HtmlRenderDocument avoided = HtmlRenderTestDriver.Render(avoidedHtml, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlDiagnostic diagnostic = Assert.Single(avoided.Diagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.ForcedFragment);

        Assert.Equal("table#oversized", diagnostic.Source);
        Assert.Contains("no safe break opportunity", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTables_FirstRowBreakBeforePromotesToTheTableFlowBoundary() {
        const string html = "<style>@page{size:180px 80px;margin:0}p,table{margin:0}"
            + "p,td{font-size:8px;line-height:12px;padding:2px}</style>"
            + "<p>BeforeTable</p><table><tbody><tr style='break-before:page'><td>FirstRow</td></tr>"
            + "<tr><td>SecondRow</td></tr></tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "BeforeTable");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FirstRow");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "FirstRow");
        Assert.Contains(rendered.Pages[1].Visuals.OfType<HtmlRenderText>(), text => text.Text == "SecondRow");
    }

    [Fact]
    public void HtmlTables_AuthoredDescendantWidthContributesToAutoTrackSizingWithoutPercentFeedback() {
        const string html = "<table style='width:360px;margin:0;table-layout:auto'><tr>"
            + "<td id='fixed-cell'><div style='width:150px'>Fixed</div></td>"
            + "<td id='percent-cell'><div style='width:100%'>Percent</div><img style='width:100%;height:1px' alt=''></td>"
            + "<td id='out-of-flow-cell'><div style='position:absolute;width:1000px'></div>Flow</td>"
            + "</tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 380D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape fixedCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#fixed-cell" && shape.Shape.StrokeWidth > 0D);
        HtmlRenderShape percentCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#percent-cell" && shape.Shape.StrokeWidth > 0D);
        HtmlRenderShape outOfFlowCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#out-of-flow-cell" && shape.Shape.StrokeWidth > 0D);

        Assert.True(fixedCell.Width > percentCell.Width * 2D);
        Assert.True(fixedCell.Width > outOfFlowCell.Width * 2D);
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateTablePaginationScene(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            if (visual is HtmlRenderSemanticGroup semantic) {
                foreach (HtmlRenderVisual child in EnumerateTablePaginationScene(semantic.Visuals)) yield return child;
            } else if (visual is HtmlRenderLayoutRegion region) {
                foreach (HtmlRenderVisual child in EnumerateTablePaginationScene(region.Visuals)) yield return child;
            } else if (visual is HtmlRenderLogicalTextGroup logical) {
                foreach (HtmlRenderVisual child in EnumerateTablePaginationScene(logical.Visuals)) yield return child;
            }
        }
    }
}
