using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfHtmlPageSelectionTests {
    [Fact]
    public void DocumentHtmlExport_AppliesPageRangesBeforeSemanticReading() {
        PdfDocument source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Page 1"));
        for (int pageNumber = 2; pageNumber <= 1001; pageNumber++) {
            int capturedPageNumber = pageNumber;
            source.PageBreak().Paragraph(paragraph => paragraph.Text("Page " + capturedPageNumber));
        }
        PdfDocument document = PdfDocument.Load(source.ToBytes());
        var options = new PdfToHtmlOptions {
            ReadOptions = new PdfReadOptions {
                Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
            },
            PageRanges = new[] { PdfPageRange.From(1001, 1001) }
        };

        PdfHtmlConversionResult result = document.ToHtmlResult(options);

        Assert.Contains("Page 1001", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1<", result.Value, StringComparison.Ordinal);
        Assert.Equal(1001, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 1001 }, result.Summary.PageNumbers);
    }

    [Fact]
    public void HtmlExport_DoesNotReportUnusedOptionalContentDefinitionsAsLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());

        PdfHtmlConversionResult result = logical.ToHtmlResult();
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;

        Assert.True(logical.HasOptionalContentGroups);
        Assert.All(logical.Pages, static page => Assert.False(page.HasOptionalContentUsage));
        Assert.Equal(logical.OptionalContentGroupCount, tableScope.OptionalContentGroupCount);
        Assert.Equal(0, tableScope.PagesWithOptionalContent);
        Assert.DoesNotContain(word.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.DoesNotContain(result.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
    }

    [Fact]
    public void VisualWordAndPowerPoint_DoNotReportUnusedOptionalContentDefinitions() {
        PdfDocument document = PdfDocument.Load(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());
        PdfWordConversionResult word = document.ToWordDocumentResult(PdfToWordOptions.CreateVisualPages());
        PdfPowerPointConversionResult powerPoint = document.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateVisualPages());
        using (word.Value)
        using (powerPoint.Value) {
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfOptionalContentGroupsFlattened");
            Assert.DoesNotContain(powerPoint.Warnings, static warning =>
                warning.Code == "PdfOptionalContentGroupsFlattened");
        }
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, true)]
    public void VisualWordAndPowerPoint_ReportOptionalContentOnlyForSelectedUsage(int pageNumber, bool expectedWarning) {
        PdfDocument document = PdfDocument.Load(PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unlayered page"))
            .PageBreak()
            .Layer("Selected layer", layer => layer.Paragraph(paragraph => paragraph.Text("Layered page")))
            .ToBytes());
        PdfToWordOptions wordOptions = PdfToWordOptions.CreateVisualPages();
        wordOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(pageNumber) };
        PdfToPowerPointOptions powerPointOptions = PdfToPowerPointOptions.CreateVisualPages();
        powerPointOptions.ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(pageNumber) };
        PdfWordConversionResult word = document.ToWordDocumentResult(wordOptions);
        PdfPowerPointConversionResult powerPoint = document.ToPowerPointPresentationResult(powerPointOptions);
        using (word.Value)
        using (powerPoint.Value) {
            Assert.Equal(expectedWarning, word.Report.Warnings.Any(static warning =>
                warning.Code == "PdfOptionalContentGroupsFlattened"));
            Assert.Equal(expectedWarning, powerPoint.Warnings.Any(static warning =>
                warning.Code == "PdfOptionalContentGroupsFlattened"));
        }
    }

    [Fact]
    public void VisualWordAndPowerPoint_ReportInconclusiveOptionalContentInspectionWithoutFailingConversion() {
        PdfDocument document = PdfDocument.Load(PdfDocument.Create()
            .Layer("Selected layer", layer => layer.Paragraph(paragraph => paragraph.Text("Layered page")))
            .ToBytes());
        PdfReadPage.OptionalContentUsageInspectionObserverForTesting = static () =>
            throw new InvalidDataException("Synthetic optional-content inspection failure.");
        try {
            PdfWordConversionResult word = document.ToWordDocumentResult(PdfToWordOptions.CreateVisualPages());
            PdfPowerPointConversionResult powerPoint = document.ToPowerPointPresentationResult(
                PdfToPowerPointOptions.CreateVisualPages());
            using (word.Value)
            using (powerPoint.Value) {
                Assert.Contains(word.Report.Warnings, static warning =>
                    warning.Code == "PdfOptionalContentUsageInspectionInconclusive");
                Assert.Contains(powerPoint.Warnings, static warning =>
                    warning.Code == "PdfOptionalContentUsageInspectionInconclusive");
            }
        } finally {
            PdfReadPage.OptionalContentUsageInspectionObserverForTesting = null;
        }
    }

    [Fact]
    public void HtmlExport_ReportsOptionalContentOnlyWhenSelectedPagesUseIt() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Unlayered page"))
            .PageBreak()
            .Layer("Selected layer", layer => layer.Paragraph(paragraph => paragraph.Text("Layered page")))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(source);

        PdfHtmlConversionResult firstPage = document.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(1, 1) }
        });
        PdfHtmlConversionResult secondPage = document.ToHtmlResult(new PdfToHtmlOptions {
            PageRanges = new[] { PdfPageRange.From(2, 2) }
        });
        PdfDocumentReadResult firstPageLogical = document.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(1)
        });
        PdfDocumentReadResult secondPageLogical = document.Read(new PdfReadOptions {
            PageSelection = PdfPageSelection.From(2)
        });
        PdfWordConversionResult firstPageWord = firstPageLogical.ToWordDocumentResult();
        PdfWordConversionResult secondPageWord = secondPageLogical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument firstWordDocument = firstPageWord.Value;
        using OfficeIMO.Word.WordDocument secondWordDocument = secondPageWord.Value;

        Assert.DoesNotContain(firstPage.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(secondPage.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened" &&
            warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Equal(0, PdfLogicalTableAnalysis.AnalyzeExtractionScope(firstPageLogical).PagesWithOptionalContent);
        Assert.Equal(1, PdfLogicalTableAnalysis.AnalyzeExtractionScope(secondPageLogical).PagesWithOptionalContent);
        Assert.DoesNotContain(firstPageWord.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(secondPageWord.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened" &&
            warning.Details["GroupCount"] == "1");
    }

    [Fact]
    public void ReverseConversions_ReportPageOptionalContentWithoutCatalogMetadata() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            BuildOptionalContentUsageWithoutCatalogMetadataPdf());

        PdfHtmlConversionResult html = logical.ToHtmlResult();
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);

        Assert.Equal(0, logical.OptionalContentGroupCount);
        Assert.True(Assert.Single(logical.Pages).HasOptionalContentUsage);
        Assert.Equal(1, tableScope.PagesWithOptionalContent);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(word.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReverseConversions_ReportOptionalContentInsideAppliedSoftMaskGroups(bool useGroupDictionaryEntry) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            BuildSoftMaskOptionalContentPdf(useGroupDictionaryEntry));

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfHtmlConversionResult html = logical.ToHtmlResult();
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;
        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation = powerPoint.Value;

        Assert.True(page.HasOptionalContentUsage);
        Assert.Equal(1, tableScope.PagesWithOptionalContent);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(word.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(powerPoint.Report.Warnings, static warning =>
            warning.Code == "PdfGroupsNotReconstructed");
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void ReverseConversions_ReportOptionalContentOnlyFromTheSelectedAnnotationAppearance(
        bool useStateDictionary,
        bool selectLayeredAppearance) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            BuildAnnotationAppearanceOptionalContentPdf(useStateDictionary, selectLayeredAppearance));

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfHtmlConversionResult html = logical.ToHtmlResult();
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;
        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation = powerPoint.Value;

        Assert.Equal(selectLayeredAppearance, page.HasOptionalContentUsage);
        Assert.Equal(selectLayeredAppearance ? 1 : 0, tableScope.PagesWithOptionalContent);
        Assert.Equal(selectLayeredAppearance, html.Report.Warnings.Any(static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened"));
        Assert.Equal(selectLayeredAppearance, word.Report.Warnings.Any(static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened"));
        Assert.Equal(selectLayeredAppearance, powerPoint.Report.Warnings.Any(static warning =>
            warning.Code == "PdfGroupsNotReconstructed"));
    }

    [Theory]
    [InlineData("A", true)]
    [InlineData("B", false)]
    public void ReverseConversions_ReportOptionalContentOnlyInsideInvokedType3Glyph(
        string layeredGlyph,
        bool expectedOptionalContent) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            BuildType3OptionalContentPdf(layeredGlyph));

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfHtmlConversionResult html = logical.ToHtmlResult();
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;
        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation = powerPoint.Value;

        Assert.Equal(expectedOptionalContent, page.HasOptionalContentUsage);
        Assert.Equal(expectedOptionalContent ? 1 : 0, tableScope.PagesWithOptionalContent);
        Assert.Equal(expectedOptionalContent, html.Report.Warnings.Any(static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened"));
        Assert.Equal(expectedOptionalContent, word.Report.Warnings.Any(static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened"));
        Assert.Equal(expectedOptionalContent, powerPoint.Report.Warnings.Any(static warning =>
            warning.Code == "PdfGroupsNotReconstructed"));
    }

    [Fact]
    public void ReverseConversions_ReportOptionalContentInsideType3GlyphSelectedByExtGStateFont() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(
            BuildType3OptionalContentPdf("A", useExtGStateFont: true));

        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfTableExtractionScopeReport tableScope = PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical);
        PdfHtmlConversionResult html = logical.ToHtmlResult();
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using OfficeIMO.Word.WordDocument wordDocument = word.Value;
        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation = powerPoint.Value;

        Assert.True(page.HasOptionalContentUsage);
        Assert.Equal(1, tableScope.PagesWithOptionalContent);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(word.Report.Warnings, static warning =>
            warning.Code == "PdfOptionalContentGroupsFlattened");
        Assert.Contains(powerPoint.Report.Warnings, static warning =>
            warning.Code == "PdfGroupsNotReconstructed");
    }

    private static byte[] BuildOptionalContentUsageWithoutCatalogMetadataPdf() {
        const string content = "/OC /UncataloguedLayer BDC BT /F1 12 Tf 20 100 Td (Layered text) Tj ET EMC";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 6 0 R >> /Properties << /UncataloguedLayer 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + content.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OCG /Name (Uncatalogued layer) >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSoftMaskOptionalContentPdf(bool useGroupDictionaryEntry) {
        const string pageContent = "/Mask gs 0 0 120 90 re f";
        string maskContent = useGroupDictionaryEntry
            ? "0 0 120 90 re f"
            : "/OC /Layer BDC 0 0 120 90 re f EMC";
        string optionalContentEntry = useGroupDictionaryEntry ? " /OC 7 0 R" : string.Empty;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /ExtGState << /Mask 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /ExtGState /SMask << /S /Alpha /G 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 120 90] /Group << /Type /Group /S /Transparency >> /Resources << /Properties << /Layer 7 0 R >> >>" + optionalContentEntry + " /Length " + maskContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            maskContent,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /OCG /Name (Mask layer) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAnnotationAppearanceOptionalContentPdf(
        bool useStateDictionary,
        bool selectLayeredAppearance) {
        const string layeredContent = "/OC /Layer BDC 0 0 40 40 re f EMC";
        const string ordinaryContent = "0 0 40 40 re f";
        string normalAppearance = useStateDictionary
            ? "<< /On 6 0 R /Off 7 0 R >>"
            : "6 0 R";
        string appearanceState = useStateDictionary
            ? " /AS /" + (selectLayeredAppearance ? "On" : "Off")
            : string.Empty;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Annots [5 0 R] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            string.Empty,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Stamp /Rect [20 20 60 60] /AP << /N " + normalAppearance + " >>" + appearanceState + " >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << /Properties << /Layer 8 0 R >> >> /Length " + layeredContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            layeredContent,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] /Length " + ordinaryContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            ordinaryContent,
            "endstream",
            "endobj",
            "8 0 obj",
            "<< /Type /OCG /Name (Annotation layer) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    [Fact]
    public void Type3OptionalContentInspectionHonorsTheGlyphInvocationLimit() {
        byte[] pdf = BuildType3OptionalContentPdf("none", pageText: "AAAA");
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 1 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            document.Pages[0].HasOptionalContentUsage());

        Assert.Equal(PdfReadLimitKind.Type3GlyphInvocations, exception.Kind);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void OptionalContentScanResolvesDirectExtGStateType3FontWithoutFontResources() {
        byte[] pdf = BuildType3OptionalContentPdf("A", useExtGStateFont: true,
            useDirectExtGStateFont: true);

        Assert.True(Assert.Single(PdfReadDocument.Open(pdf).Pages).HasOptionalContentUsage());
    }

    private static byte[] BuildType3OptionalContentPdf(string layeredGlyph, bool useExtGStateFont = false,
        string pageText = "A", bool useDirectExtGStateFont = false) {
        string pageContent = useExtGStateFont
            ? "BT /FontState gs 20 100 Td (" + pageText + ") Tj ET"
            : "BT /F3 24 Tf 20 100 Td (" + pageText + ") Tj ET";
        string extGState = useExtGStateFont
            ? " /ExtGState << /FontState << /Font [" + (useDirectExtGStateFont ? "5 0 R" : "/F3") + " 24] >> >>"
            : string.Empty;
        string fontResources = useDirectExtGStateFont ? string.Empty : "/Font << /F3 5 0 R >>";
        const string ordinaryGlyph = "0 0 500 700 d1 0 0 500 700 re f";
        const string layeredGlyphContent = "0 0 500 700 d1 /OC /Layer BDC 0 0 500 700 re f EMC";
        string glyphA = string.Equals(layeredGlyph, "A", StringComparison.Ordinal) ? layeredGlyphContent : ordinaryGlyph;
        string glyphB = string.Equals(layeredGlyph, "B", StringComparison.Ordinal) ? layeredGlyphContent : ordinaryGlyph;
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << " + fontResources + extGState + " >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + pageContent.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            pageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R /B 7 0 R >> /Encoding << /Type /Encoding /Differences [65 /A /B] >> /FirstChar 65 /LastChar 66 /Widths [500 500] /Resources << /Properties << /Layer 8 0 R >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Length " + glyphA.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            glyphA,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Length " + glyphB.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            glyphB,
            "endstream",
            "endobj",
            "8 0 obj",
            "<< /Type /OCG /Name (Glyph layer) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }
}
