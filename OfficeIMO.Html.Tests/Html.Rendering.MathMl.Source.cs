using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string SourceMath = "<math xmlns='http://www.w3.org/1998/Math/MathML' alttext='x squared'>\r\n<!--source--><msup><mi><![CDATA[x]]></mi><mn>2</mn></msup></math>";

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void HtmlMathMlSource_RecoversExactMarkupAcrossNativeAndOwnedSnapshots(int snapshot) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<!doctype html><p id='other'>BEFORE " + SourceMath + " AFTER</p>");
        if (snapshot == 1) document = HtmlConversionDocument.FromDocument(document.Document.Clone());
        if (snapshot == 2) document = document.Edit(tree => tree.QuerySelector("#other")!.SetAttribute("title", "unrelated edit"));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document);
        HtmlRenderSemanticGroup formula = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Formula);
        Assert.True(formula.MathMlSource!.IsOriginalMarkup, formula.MathMlSource.MathMl);
        Assert.Equal(SourceMath, formula.MathMlSource.MathMl);
        byte[] pdf = document.ToPdfBytes();
        PdfCore.PdfExtractedAttachment attachment = Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf));
        Assert.Equal(Encoding.UTF8.GetBytes(SourceMath), attachment.Bytes);
        Assert.Equal("application/mathml+xml", attachment.MimeType);
        Assert.Equal(PdfCore.PdfAssociatedFileRelationship.Supplement, attachment.Relationship);
        Assert.StartsWith("%PDF-2.0", Encoding.ASCII.GetString(pdf, 0, 8));
        Assert.Contains("BEFORE", PdfCore.PdfReadDocument.Open(pdf).ExtractText());
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
    }

    [Fact]
    public void HtmlMathMlSource_ContextualFragmentAndImportedNodeRetainTheirOwnSource() {
        var parser = OfficeIMO.Html.Providers.AngleSharpHtmlParser.Instance;
        var tree = parser.ParseDocument("<body><p>BEFORE</p></body>", new OfficeIMO.Html.Dom.HtmlParseOptions()).Clone();
        var body = tree.QuerySelector("body")!;
        var fragment = parser.ParseFragment(SourceMath, body, new OfficeIMO.Html.Dom.HtmlParseOptions());
        body.AppendChild(tree.ImportNode(fragment));
        byte[] pdf = HtmlConversionDocument.FromDocument(tree).ToPdfBytes();
        Assert.Equal(SourceMath, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf)).Bytes));
    }

    [Fact]
    public void HtmlMathMlSource_EditedFormulaNeverAssociatesStaleOriginal() {
        HtmlConversionDocument original = HtmlConversionDocument.Parse("<p>BEFORE " + SourceMath + " AFTER</p>");
        HtmlConversionDocument edited = original.Edit(tree => tree.QuerySelector("mi")!.TextContent = "z");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(edited);
        HtmlMathMlSource source = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Formula).MathMlSource!;
        Assert.False(source.IsOriginalMarkup);
        Assert.Equal("z", XElement.Parse(source.MathMl).Descendants().Single(element => element.Name.LocalName == "mi").Value);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlSourceNormalized);
        Assert.Equal(source.MathMl, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(edited.ToPdfBytes())).Bytes));
        Assert.Equal(SourceMath, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(original.ToPdfBytes())).Bytes));
    }

    [Fact]
    public void HtmlMathMlSource_StaleOriginalComparisonLimitKeepsTheValidEditedFormulaSource() {
        string originalMarkup = "<math xmlns='http://www.w3.org/1998/Math/MathML'>"
            + string.Concat(Enumerable.Repeat("<mrow>", 130)) + "<mi>x</mi>"
            + string.Concat(Enumerable.Repeat("</mrow>", 130)) + "</math>";
        HtmlConversionDocument edited = HtmlConversionDocument.Parse(originalMarkup).Edit(tree => {
            var math = tree.QuerySelector("math")!;
            math.TextContent = string.Empty;
            var identifier = tree.CreateElement("mi", math.NamespaceUri);
            identifier.TextContent = "z";
            math.AppendChild(identifier);
        });
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(edited);
        HtmlRenderSemanticGroup formula = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Formula);
        Assert.NotNull(formula.MathMlSource);
        Assert.False(formula.MathMlSource.IsOriginalMarkup);
        Assert.Equal("z", XElement.Parse(formula.MathMlSource.MathMl).Value);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlSourceNormalized);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlSourceUnavailable);
        Assert.Equal(formula.MathMlSource.MathMl, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(edited.ToPdfBytes())).Bytes));
    }

    [Theory]
    [InlineData("<math alttext=x><mi>&alpha;</mi></math>", "α")]
    [InlineData("<math><mi>x</mi><p>outside</p></math>", "x")]
    public void HtmlMathMlSource_RecoveredHtmlUsesCurrentStandaloneMathMl(string math, string expected) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<p>BEFORE</p>" + math + "<p>AFTER</p>");
        HtmlMathMlSource source = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Formula).MathMlSource!;
        Assert.False(source.IsOriginalMarkup);
        XElement xml = XElement.Parse(source.MathMl);
        Assert.Equal("http://www.w3.org/1998/Math/MathML", xml.Name.NamespaceName);
        Assert.Equal(expected, xml.Descendants().Single(element => element.Name.LocalName == "mi").Value);
        Assert.DoesNotContain("outside", source.MathMl);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlSourceNormalized);
    }

    [Fact]
    public void HtmlMathMlSource_TokenCaptureIgnoresCommentsScriptAndQuotedClosingTags() {
        string math = "<math xmlns='http://www.w3.org/1998/Math/MathML' alttext='literal &lt;/math>'><mi>x</mi></math>";
        string html = "<!--<math><mi>fake</mi></math>--><script>const x = '<math><mi>fake</mi></math>';</script><p>BEFORE " + math + " AFTER</p>";
        Assert.Equal(math, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(HtmlConversionDocument.Parse(html).ToPdfBytes())).Bytes));
    }

    [Fact]
    public void HtmlMathMlSource_ValidRootAfterForeignContentRecoveryStillRetainsExactMarkup() {
        string html = "<math><mi>first</mi><p>outside</p>" + SourceMath + "<p>AFTER</p>";
        var attachments = PdfCore.PdfAttachmentExtractor.ExtractAttachments(HtmlConversionDocument.Parse(html).ToPdfBytes());
        Assert.Contains(attachments, file => Encoding.UTF8.GetString(file.Bytes) == SourceMath);
        Assert.All(attachments, file => Assert.DoesNotContain("outside", Encoding.UTF8.GetString(file.Bytes)));
    }

    [Fact]
    public void HtmlMathMlSource_IdenticalVisibleFormulasShareSupplementWhileHiddenSourceIsAbsent() {
        string html = "<p>BEFORE " + SourceMath + " BETWEEN " + SourceMath + " AFTER</p><math style='display:none'><mi>hidden</mi></math>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        Assert.Equal(SourceMath, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf)).Bytes));
        PdfCore.PdfRawDocumentView raw = PdfCore.PdfReadDocument.Open(pdf).RawStructure();
        var formulas = raw.Objects.Where(obj => obj.Value.Entries.TryGetValue("S", out var role) && role.Text == "Formula").ToArray();
        Assert.Equal(2, formulas.Length);
        Assert.Equal(Assert.Single(formulas[0].Value.Entries["AF"].Items).ReferenceObjectNumber, Assert.Single(formulas[1].Value.Entries["AF"].Items).ReferenceObjectNumber);
    }

    [Theory]
    [InlineData("style='transform:translate(8px,4px)'", 4D)]
    [InlineData("dir='rtl'", 4D)]
    [InlineData("style='writing-mode:vertical-rl'", 4D)]
    [InlineData("style='margin-top:140px'", 1D)]
    public void HtmlMathMlSource_PaintTransformsAndPaginationKeepRecoverableSource(string attributes, double pageHeight) {
        string html = "<html lang='en'><body style='margin:0'><div " + attributes + ">BEFORE " + SourceMath + " AFTER</div></body></html>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { PageSize = new OfficeIMO.Drawing.OfficePageSize(3D, pageHeight), HonorCssPageRules = false, Margins = HtmlRenderMargins.All(0D) });
        Assert.Equal(SourceMath, Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf)).Bytes));
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        Assert.Equal("x squared", Assert.Single(tagged.StructureElements, item => item.StructureType == "Formula").AlternateText);
    }

    [Fact]
    public void HtmlMathMlSource_UntaggedOutputDoesNotLeaveOrphanPayload() {
        var options = new HtmlToPdfOptions();
        options.PdfOptions.TaggedStructureMode = PdfCore.PdfTaggedStructureMode.None;
        byte[] pdf = HtmlConversionDocument.Parse(SourceMath).ToPdfBytes(options);
        Assert.Empty(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf));
        Assert.DoesNotContain("%PDF-2.0", Encoding.ASCII.GetString(pdf, 0, 8));
        Assert.Contains("x^(2)", PdfCore.PdfReadDocument.Open(pdf).ExtractText());
    }

    [Fact]
    public void HtmlMathMlSource_ExplicitDateSurvivesAdapterOptionSnapshots() {
        var date = new DateTimeOffset(2026, 9, 30, 12, 0, 0, TimeSpan.Zero);
        var options = new HtmlToPdfOptions { MathMlSourceModificationDate = date };
        Assert.Equal(date, options.ClonePdf().MathMlSourceModificationDate);
        byte[] pdf = HtmlConversionDocument.Parse(SourceMath).ToPdfBytes(options);
        Assert.Equal(date, Assert.Single(PdfCore.PdfInspector.Inspect(pdf).Attachments).ModificationDate);
    }
}
