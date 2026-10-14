using System.Threading;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioTextInheritanceBoundaryTests {
    private static readonly XNamespace Native = VisioDocument.VisioNamespace;

    [Fact]
    public void ReplacingTextUsesDefaultRowsWithoutRevivingOldMarkersAfterReopen() {
        VisioDocument document = Load("<StyleSheet ID='6'><Char IX='0'><Size>0.1666666666666667</Size></Char></StyleSheet>",
            "<Char IX='0'><Size>0.25</Size></Char><Para IX='0'><Bullet>1</Bullet></Para><Para IX='9'><Bullet>7</Bullet></Para>",
            "<cp IX='0'/><pp IX='9'/>Original");
        document.Pages[0].Shapes[0].Text = "Replacement";
        foreach (VisioDocument candidate in Candidates(document)) {
            VisioRichTextProjection projection = VisioRichTextProjection.Create(candidate.Pages[0], candidate.Pages[0].Shapes[0], 72, default)!;
            OfficeRichTextParagraph paragraph = Assert.Single(projection.Paragraphs);
            Assert.Equal("•", paragraph.Label!.Run.Text);
            OfficeRichTextRun run = Assert.Single(paragraph.Runs);
            Assert.Equal("Replacement", run.Text);
            Assert.Equal(18, run.FontSize);
            XElement svg = XElement.Parse(candidate.Pages[0].ToSvg());
            XElement[] painted = svg.Descendants().Where(e => e.Name.LocalName == "text" && !string.IsNullOrWhiteSpace(e.Value)).ToArray();
            Assert.Equal(new[] { "•", "Replacement" }, painted.Select(e => e.Value));
            Assert.All(painted, e => Assert.Equal("24", (string?)e.Attribute("font-size")));
        }
    }

    [Fact]
    public void ImplicitAndPaddedRowIndicesKeepDistinctRunFormatting() {
        VisioDocument document = Load("<StyleSheet ID='6'><Char IX='0'><Size>0.1666666666666667</Size></Char></StyleSheet>",
            "<Char><Color>#cc1122</Color></Char><Char><Color>#2244cc</Color></Char>",
            "<cp IX='00'/>RED<cp IX='01'/>BLUE");
        foreach (VisioDocument candidate in Candidates(document)) {
            VisioShape shape = candidate.Pages[0].Shapes[0];
            VisioRichTextProjection projection = VisioRichTextProjection.Create(candidate.Pages[0], shape, 72, default)!;
            Assert.Equal(OfficeColor.FromRgb(204, 17, 34), projection.Runs[0].Color);
            Assert.Equal(OfficeColor.FromRgb(34, 68, 204), projection.Runs[1].Color);
        }
    }

    [Fact]
    public void DisabledStyleDoesNotLeakOwnTextPropertiesAndReportsUnqualifiedTraversal() {
        VisioDocument document = Load("<StyleSheet ID='6' TextStyle='7'><StyleProp><EnableTextProps>0</EnableTextProps></StyleProp>"
            + "<Char IX='0'><Size>0.5</Size></Char></StyleSheet><StyleSheet ID='7'><Char IX='0'><Size>0.25</Size></Char></StyleSheet>");
        foreach (VisioDocument candidate in Candidates(document)) {
            OfficeImageExportResult result = candidate.Pages[0].ExportImage(OfficeImageExportFormat.Svg);
            Assert.Contains(result.Diagnostics, d => d.Code == VisioNativeTextStyleResolver.DiagnosticCode && d.Message.Contains("excludes"));
            XElement svg = XElement.Parse(candidate.Pages[0].ToSvg());
            Assert.DoesNotContain(svg.Descendants(), e => (string?)e.Attribute("font-size") == "48");
        }
    }

    [Theory]
    [InlineData("cycle")]
    [InlineData("missing")]
    [InlineData("depth")]
    public void IncompleteStyleChainsKeepCachedValuesAndReportLoss(string kind) {
        VisioDocument document = Load("<StyleSheet ID='6' TextStyle='7'><Char IX='0'><Size>0.1666666666666667</Size></Char></StyleSheet>"
            + (kind == "cycle" ? "<StyleSheet ID='7' TextStyle='6'/>" : kind == "depth" ? string.Concat(Enumerable.Range(7, 70)
                .Select(id => $"<StyleSheet ID='{id}' TextStyle='{id + 1}'/>")) : ""));
        string source = Encoding.UTF8.GetString(document.ToLegacyXmlResult().Value);
        foreach (VisioDocument candidate in Candidates(document)) {
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            var resolver = new VisioNativeTextStyleResolver(candidate, default, diagnostics, "test");
            VisioRichTextProjection projection = VisioRichTextProjection.Create(candidate.Pages[0], candidate.Pages[0].Shapes[0], 72, default, resolver)!;
            Assert.Equal(12, Assert.Single(projection.Runs).FontSize, 8);
            Assert.Single(diagnostics, d => d.Code == VisioNativeTextStyleResolver.DiagnosticCode);
            foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Svg, OfficeImageExportFormat.Png }) {
                OfficeImageExportResult result = candidate.Pages[0].ExportImage(format, new VisioImageExportOptions { Supersampling = 1 });
                Assert.Contains(result.Diagnostics, d => d.Code == VisioNativeTextStyleResolver.DiagnosticCode && d.LossKind == OfficeConversionLossKind.Approximation);
                OfficeImageExportPolicyException error = Assert.Throws<OfficeImageExportPolicyException>(() => candidate.Pages[0].ExportImage(format,
                    new VisioImageExportOptions { Supersampling = 1, Policy = new OfficeImageExportPolicy { RequireNoLoss = true } }));
                Assert.Contains(error.Diagnostics, d => d.Code == VisioNativeTextStyleResolver.DiagnosticCode);
            }
        }
        Assert.Equal(source, Encoding.UTF8.GetString(document.ToLegacyXmlResult().Value));
    }

    [Fact]
    public void DeletedRowsAndSectionsDoNotResurrectInheritedProperties() {
        VisioDocument original = Load("<StyleSheet ID='6'><Char IX='0'><Size>0.25</Size><Color>#cc1122</Color></Char></StyleSheet>", "<Char IX='0' Del='1'/>");
        foreach (bool sectionDeleted in new[] { false, true }) {
            VisioDocument document = sectionDeleted ? WithSectionDeletion(original) : original;
            foreach (VisioDocument candidate in sectionDeleted ? new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) } : Candidates(document)) {
                OfficeRichTextRun run = Assert.Single(VisioRichTextProjection.Create(candidate.Pages[0], candidate.Pages[0].Shapes[0], 72, default)!.Runs);
                Assert.Equal(10, run.FontSize);
                Assert.Equal(OfficeColor.FromRgb(17, 24, 39), run.Color);
            }
            if (sectionDeleted) {
                VisioXmlConversionReport report = document.ToLegacyXmlResult().Report;
                Assert.Contains(report.FidelityDiagnostics, d => d.Code == "VDX_SECTION_DELETION");
                Assert.Throws<OfficeConversionException>(() => report.RequireNoLoss());
            }
        }
    }

    private static VisioDocument WithSectionDeletion(VisioDocument document) {
        using var stream = new MemoryStream();
        byte[] bytes = document.ToBytes();
        stream.Write(bytes, 0, bytes.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            ZipArchiveEntry entry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            XElement section = xml.Descendants(Native + "Section").Single(e => (string?)e.Attribute("N") == "Character");
            section.SetAttributeValue("Del", "1");
            section.Element(Native + "Row")!.Attribute("Del")!.Remove();
            entry.Delete();
            using var output = archive.CreateEntry("visio/pages/page1.xml").Open();
            xml.Save(output);
        }
        return VisioDocument.Load(new MemoryStream(stream.ToArray()));
    }

    [Theory]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void DeletedNonzeroTextSectionsDoNotReviveModeledCaches(bool connector, bool character, bool unsupportedParagraph) {
        VisioDocument original = Load("<StyleSheet ID='6'><Char IX='0'><Size>0.5</Size></Char>"
            + "<Para IX='0'><HorzAlign>0</HorzAlign></Para></StyleSheet>",
            "<Char IX='9'><Font>4</Font><Size>0.25</Size></Char><Para IX='9'>"
                + (unsupportedParagraph ? "<IndLeft>-0.1</IndLeft>" : "") + "<HorzAlign>2</HorzAlign></Para>",
            "<cp IX='9'/><pp IX='9'/>Deleted", connector);
        original.PreservedFaceNamesElements.Add(new XElement(Native + "FaceName", new XAttribute("ID", "4"),
            new XAttribute("Name", "Deleted cached font")));
        VisioDocument document = WithNonzeroSectionDeletion(original, character);
        foreach (VisioDocument candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            byte[] before = candidate.ToLegacyXmlResult().Value;
            VisioPage page = candidate.Pages[0];
            VisioRichTextProjection projection = Assert.IsType<VisioRichTextProjection>(connector
                ? VisioRichTextProjection.Create(page, Assert.Single(page.Connectors), 72, default)
                : VisioRichTextProjection.Create(page, Assert.Single(page.Shapes), 72, default));
            if (unsupportedParagraph) Assert.Empty(projection.Paragraphs);
            else Assert.Equal(character ? OfficeTextAlignment.Right : OfficeTextAlignment.Center, Assert.Single(projection.Paragraphs).Alignment);
            Assert.Equal(character ? connector ? 9D : 10D : 18D, Assert.Single(projection.Runs).FontSize, 8);
            if (character) Assert.DoesNotContain("Deleted cached font", Assert.Single(projection.Runs).FontFamily);
            XElement painted = Assert.Single(XElement.Parse(page.ToSvg()).Descendants(), e => e.Name.LocalName == "text");
            double points = double.Parse((string)painted.Attribute("font-size")!, System.Globalization.CultureInfo.InvariantCulture) * 72 / 96;
            Assert.InRange(points, (character ? connector ? 9D : 10D : 18D) - .001, (character ? connector ? 9D : 10D : 18D) + .001);
            _ = page.ToPng(new VisioPngSaveOptions { Supersampling = 1 });
            if (character) Assert.DoesNotContain(page.ExportImage(OfficeImageExportFormat.Svg).Diagnostics,
                diagnostic => diagnostic.Message.Contains("Deleted cached font"));
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
            using var archive = new ZipArchive(new MemoryStream(candidate.ToBytes()), ZipArchiveMode.Read);
            using var source = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XElement section = XDocument.Load(source).Descendants(Native + "Section")
                .Single(e => (string?)e.Attribute("N") == (character ? "Character" : "Paragraph"));
            Assert.Equal("1", (string?)section.Attribute("Del"));
            Assert.Equal("9", (string?)Assert.Single(section.Elements(Native + "Row")).Attribute("IX"));
        }
    }

    private static VisioDocument WithNonzeroSectionDeletion(VisioDocument document, bool character) {
        using var stream = new MemoryStream();
        byte[] bytes = document.ToBytes();
        stream.Write(bytes, 0, bytes.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            ZipArchiveEntry entry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            xml.Descendants(Native + "Section").Single(e => (string?)e.Attribute("N") == (character ? "Character" : "Paragraph"))
                .SetAttributeValue("Del", "1");
            entry.Delete();
            using var output = archive.CreateEntry("visio/pages/page1.xml").Open();
            xml.Save(output, SaveOptions.DisableFormatting);
        }
        return VisioDocument.Load(new MemoryStream(stream.ToArray()));
    }

    [Fact]
    public void ThemedAndUncachedCellsArePreservedWithExplicitFallbackDiagnostics() {
        VisioDocument document = Load("<StyleSheet ID='6' TextStyle='7'/><StyleSheet ID='7'><Char IX='0'><Size>0.25</Size></Char></StyleSheet>");
        XElement style = document.PreservedAdditionalStyleSheets.Single(e => (string?)e.Attribute("ID") == "6");
        style.Add(new XElement(Native + "Section", new XAttribute("N", "Character"), new XElement(Native + "Row", new XAttribute("IX", "0"),
            new XElement(Native + "Cell", new XAttribute("N", "Font"), new XAttribute("V", "Themed")),
            new XElement(Native + "Cell", new XAttribute("N", "Color"), new XAttribute("V", "Themed")),
            new XElement(Native + "Cell", new XAttribute("N", "Size"), new XAttribute("F", "UNSUPPORTED()")))));
        string source = style.ToString();
        foreach (VisioDocument candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            var resolver = new VisioNativeTextStyleResolver(candidate, default, diagnostics);
            OfficeRichTextRun run = Assert.Single(VisioRichTextProjection.Create(candidate.Pages[0], candidate.Pages[0].Shapes[0], 72, default, resolver)!.Runs);
            Assert.Equal(18, run.FontSize, 8);
            Assert.Equal(OfficeColor.FromRgb(17, 24, 39), run.Color);
            Assert.DoesNotContain("themed", run.FontFamily.ToLowerInvariant());
            Assert.Contains(diagnostics, d => d.Message.Contains("theme"));
            Assert.Contains(diagnostics, d => d.Message.Contains("uncached"));
        }
        Assert.Equal(source, style.ToString());
    }

    [Fact]
    public void MissingBindingDoesNotSelectTheNextDrawnShapeDefaultAndResolutionCancels() {
        VisioDocument document = Load("<StyleSheet ID='6'><Char IX='0'><Size>0.25</Size></Char></StyleSheet>");
        VisioShape shape = document.Pages[0].Shapes[0];
        shape.NativeStyleReferences = VisioNativeStyleReferences.Read(new XElement(Native + "Shape"));
        document.PreservedDocumentSettingsAttributes.Add(new XAttribute("DefaultTextStyle", "6"));
        foreach (VisioDocument candidate in Candidates(document)) {
            XElement svg = XElement.Parse(candidate.Pages[0].ToSvg());
            Assert.Equal("13.333", (string?)Assert.Single(svg.Descendants(), e => e.Name.LocalName == "text").Attribute("font-size"));
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => new VisioNativeTextStyleResolver(document, cancelled.Token));
    }

    private static VisioDocument Load(string styles, string local = "", string text = "VISIBLE", bool connector = false) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><StyleSheets>" + styles + "</StyleSheets>"
        + "<Pages><Page ID='0'><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>2</PageHeight></PageProps></PageSheet><Shapes>"
        + "<Shape ID='1' TextStyle='6'>" + (connector
            ? "<XForm1D><BeginX>0.5</BeginX><BeginY>1</BeginY><EndX>3.5</EndX><EndY>1</EndY></XForm1D>"
                + "<TextXForm><TxtWidth>3</TxtWidth><TxtHeight>1</TxtHeight></TextXForm>"
            : "<XForm><PinX>2</PinX><PinY>1</PinY><Width>3</Width><Height>1</Height></XForm>") + local + "<Text>" + text + "</Text></Shape>"
        + "</Shapes></Page></Pages></VisioDocument>"))).Value;

    private static IEnumerable<VisioDocument> Candidates(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }
}
