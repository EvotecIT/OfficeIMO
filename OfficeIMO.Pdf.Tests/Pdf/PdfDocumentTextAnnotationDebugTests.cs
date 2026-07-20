using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void TextAnnotations_RenderFromFlowAndCanvas() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false,
                Debug = new PdfDebugOptions {
                    ShowFlowObjectBoxes = true,
                    ShowCanvasItemBoxes = true
                }
            })
            .TextAnnotation(
                "Review flow anchor",
                width: 24,
                height: 24,
                align: PdfAlign.Right,
                icon: PdfTextAnnotationIcon.Key,
                color: new PdfColor(1D, 0D, 0D),
                open: true)
            .FreeTextAnnotation(
                "Flow free text wraps across several words for reviewers",
                width: 140,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D),
                textAlign: PdfAlign.Center,
                padding: 4D,
                lineHeight: 11D)
            .HighlightAnnotation("Flow highlight", width: 120, height: 12, color: new PdfColor(1D, 0.9D, 0.1D))
            .Canvas(canvas => canvas
                .Text("Canvas text", 72, 120, 120, 28)
                .TextAnnotation("Canvas note", 200, 120, 20, 20, PdfTextAnnotationIcon.Help, new PdfColor(0D, 0.5D, 1D))
                .FreeTextAnnotation("Canvas free text wraps as a right aligned reviewer callout", 72, 170, 150, 48, borderColor: new PdfColor(0.2D, 0.4D, 0.8D), fillColor: new PdfColor(0.95D, 0.98D, 1D), textAlign: PdfAlign.Right, padding: 5D, lineHeight: 12D)
                .HighlightAnnotation("Canvas highlight", 72, 220, 120, 14, new PdfColor(1D, 0.9D, 0.1D)))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Annots [", pdf, StringComparison.Ordinal);
        Assert.Equal(2, CountTextAnnotationOccurrences(pdf));
        Assert.Equal(2, CountOccurrences(pdf, "/Subtype /FreeText"));
        Assert.Equal(2, CountOccurrences(pdf, "/Subtype /Highlight"));
        Assert.Contains("/Name /Key", pdf, StringComparison.Ordinal);
        Assert.Contains("/Name /Help", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <52657669657720666C6F7720616E63686F72>", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <43616E766173206E6F7465>", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <466C6F7720667265652074657874207772617073206163726F7373207365766572616C20776F72647320666F7220726576696577657273>", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <43616E76617320667265652074657874207772617073206173206120726967687420616C69676E65642072657669657765722063616C6C6F7574>", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <466C6F7720686967686C69676874>", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents <43616E76617320686967686C69676874>", pdf, StringComparison.Ordinal);
        Assert.Contains("/C [1 0 0]", pdf, StringComparison.Ordinal);
        Assert.Contains("/C [0.2 0.4 0.8]", pdf, StringComparison.Ordinal);
        Assert.Contains("/IC [0.95 0.98 1]", pdf, StringComparison.Ordinal);
        Assert.Contains("/QuadPoints [", pdf, StringComparison.Ordinal);
        Assert.Equal(4, CountOccurrences(pdf, "/AP << /N "));
        Assert.Equal(2, CountOccurrences(pdf, "/OfficeIMOHighlightGs gs"));
        Assert.Equal(2, CountOccurrences(pdf, "/BM /Multiply"));
        Assert.Equal(2, CountOccurrences(pdf, "/ca 0.35"));
        Assert.True(CountOccurrences(pdf, "BT\n/Helv") >= 4);
        Assert.Contains("/Subtype /Form /BBox [0 0 140 44]", pdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form /BBox [0 0 120 12]", pdf, StringComparison.Ordinal);
        Assert.Contains("/Open true", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 1 RG", pdf, StringComparison.Ordinal);
        Assert.Contains("0 0.65 1 RG", pdf, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.True(info.HasAnnotations);
        Assert.Equal(6, info.AnnotationCount);
        Assert.Equal(6, info.Pages[0].Annotations.Count);
        Assert.Equal(2, info.GetAnnotationsBySubtype("Text").Count);
        Assert.Equal(2, info.GetAnnotationsBySubtype("FreeText").Count);
        Assert.Equal(2, info.GetAnnotationsBySubtype("Highlight").Count);
        Assert.All(info.GetAnnotationsBySubtype("FreeText"), annotation => Assert.True(annotation.HasNormalAppearance));
        Assert.All(info.GetAnnotationsBySubtype("Highlight"), annotation => Assert.True(annotation.HasNormalAppearance));
    }

    [Fact]
    public void TextAnnotations_EncodeContentsAsPdfTextStrings() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .TextAnnotation("Zażółć 漢字", width: 24, height: 24)
            .FreeTextAnnotation("Komentarz 漢字", width: 140, height: 44)
            .HighlightAnnotation("Ważne 漢字", width: 120, height: 14)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Contains("/Contents <FEFF", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("Zażółć", pdf, StringComparison.Ordinal);
        Assert.Contains(info.GetAnnotationsBySubtype("Text"), annotation => annotation.Contents == "Zażółć 漢字");
        Assert.Contains(info.GetAnnotationsBySubtype("FreeText"), annotation => annotation.Contents == "Komentarz 漢字");
        Assert.Contains(info.GetAnnotationsBySubtype("Highlight"), annotation => annotation.Contents == "Ważne 漢字");
    }

    [Fact]
    public void VisualAnnotations_FlattenIntoPageContentWhenRequested() {
        PdfOptions options = new PdfOptions {
            CompressContentStreams = false,
            FlattenVisualAnnotations = true
        };

        byte[] bytes = PdfDocument.Create(options.Clone())
            .FreeTextAnnotation(
                "Flattened reviewer note wraps into form content",
                width: 140,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D),
                textAlign: PdfAlign.Center,
                padding: 4D,
                lineHeight: 11D)
            .HighlightAnnotation("Flattened highlight", width: 120, height: 14, color: new PdfColor(1D, 0.9D, 0.1D))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.True(options.Clone().FlattenVisualAnnotations);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/XObject << /Ann1 ", pdf, StringComparison.Ordinal);
        Assert.Contains("/Ann1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/Ann2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Form /BBox [0 0 140 44]", pdf, StringComparison.Ordinal);
        Assert.Contains("BT\n/Helv", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/BM /Multiply", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.35", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.9 0.1 rg 0 0 120 14 re f", pdf, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.False(info.HasAnnotations);
        Assert.Equal(0, info.AnnotationCount);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenFromAppearanceStreams() {
        byte[] annotated = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .FreeTextAnnotation(
                "Existing free text appearance",
                width: 150,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D),
                textAlign: PdfAlign.Center,
                padding: 4D,
                lineHeight: 11D)
            .HighlightAnnotation("Existing highlight appearance", width: 120, height: 14, color: new PdfColor(1D, 0.9D, 0.1D))
            .ToBytes();

        PdfDocumentInfo before = PdfInspector.Inspect(annotated);
        Assert.Equal(2, before.AnnotationCount);
        Assert.Equal(2, CountOccurrences(Encoding.ASCII.GetString(annotated), "/AP << /N "));

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo after = PdfInspector.Inspect(flattened);

        Assert.False(after.HasAnnotations);
        Assert.Equal(0, after.AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/XObject << /OfficeIMOAnnot1 ", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("BT\n/Helv", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/BM /Multiply", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.35", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.9 0.1 rg 0 0 120 14 re f", pdf, StringComparison.Ordinal);

        using var input = new MemoryStream(annotated);
        byte[] flattenedFromStream = PdfAnnotationFlattener.FlattenVisualAnnotations(input);
        Assert.Equal(0, PdfInspector.Inspect(flattenedFromStream).AnnotationCount);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenFromAppearanceStateDictionaries() {
        byte[] annotated = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .FreeTextAnnotation(
                "State dictionary appearance",
                width: 150,
                height: 44,
                borderColor: new PdfColor(0.2D, 0.4D, 0.8D),
                fillColor: new PdfColor(0.95D, 0.98D, 1D))
            .ToBytes();

        annotated = ConvertFirstNormalAppearanceToStateDictionary(annotated);
        Assert.Equal(1, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("/XObject << /OfficeIMOAnnot1 ", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("BT\n/Helv", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesMissingAppearanceStreams() {
        byte[] annotated = BuildVisualAnnotationPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo after = PdfInspector.Inspect(flattened);

        Assert.False(after.HasAnnotations);
        Assert.Equal(0, after.AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/Font << /Helv << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding", pdf, StringComparison.Ordinal);
        Assert.Contains("0.95 0.98 1 rg 0 0 150 44 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 RG 1 w", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf 0.1 0.2 0.3 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/BM /Multiply", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.35", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.9 0.1 rg 0 0 120 14 re f", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void LoadedPdfImageExport_SynthesizesSupportedMissingAnnotationAppearancesWithPolicyEvidence() {
        byte[] annotated = BuildVisualAnnotationPdfWithoutAppearances();
        PdfReadDocument document = PdfReadDocument.Open(annotated);
        PdfReadPage page = document.Pages[0];

        OfficeDrawing drawing = page.ToDrawing();
        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilities =
            page.GetRenderCapabilityDiagnostics();
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        Assert.Contains(
            drawing.Elements.OfType<OfficeDrawingText>(),
            item => item.Text.Contains("Synthetic free text", StringComparison.Ordinal));
        Assert.Contains(
            drawing.Shapes,
            item => item.Shape.FillColor is OfficeColor fill && fill.A > 0);
        Assert.Equal(
            2,
            capabilities.Count(item =>
                item.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId));
        Assert.DoesNotContain(
            capabilities,
            item => item.Code == PdfRenderCapabilities.AnnotationAppearanceId);
        Assert.Contains(
            result.Diagnostics,
            item =>
                item.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId &&
                item.LossKind == OfficeImageExportLossKind.Approximation);

        var strict = new PdfImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };
        OfficeImageExportPolicyException exception =
            Assert.Throws<OfficeImageExportPolicyException>(
                () => page.ExportImage(OfficeImageExportFormat.Png, strict));
        Assert.Contains(
            exception.Diagnostics,
            item => item.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId);
    }

    [Fact]
    public void LoadedPdfImageExport_StillReportsUnsupportedMissingAnnotationAppearances() {
        byte[] annotated = PdfDocument.Create()
            .TextAnnotation("Unsupported icon annotation")
            .Paragraph(paragraph => paragraph.Text("Page content"))
            .ToBytes();
        PdfReadPage page = PdfReadDocument.Open(annotated).Pages[0];

        IReadOnlyList<PdfRenderCapabilityDiagnostic> capabilities =
            page.GetRenderCapabilityDiagnostics();

        Assert.Contains(
            capabilities,
            item =>
                item.Code == PdfRenderCapabilities.AnnotationAppearanceId &&
                item.Subject?.StartsWith("Text[", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(
            capabilities,
            item => item.Code == PdfRenderCapabilities.SynthesizedAnnotationAppearanceId);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextGrayDefaultAppearance() {
        byte[] annotated = BuildFreeTextDefaultAppearanceColorPdfWithoutAppearance(
            "/Helv 14 Tf 0.35 g",
            "Gray default appearance");
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/DA (/Helv 14 Tf 0.35 g)", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 14 Tf 0.35 0.35 0.35 rg", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextCmykDefaultAppearance() {
        byte[] annotated = BuildFreeTextDefaultAppearanceColorPdfWithoutAppearance(
            "/Helv 11 Tf 0.1 0.2 0.3 0.25 k",
            "CMYK default appearance");
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/DA (/Helv 11 Tf 0.1 0.2 0.3 0.25 k)", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 11 Tf 0.675 0.6 0.525 rg", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextDefaultStyle() {
        byte[] annotated = BuildFreeTextDefaultStylePdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/DA", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/DS (font: Helvetica 15pt; color: #336699; text-align: right)", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 15 Tf 0.2 0.4 0.6 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("<5374796C65642044532074657874> Tj", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("BT /Helv 10 Tf 0 0 0 rg", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenRendersRichContentsSpanAppearance() {
        byte[] annotated = BuildFreeTextRichContentsPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Contents (", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/RC (<body><p>Rich <b>text</b><br/>second &amp; line</p></body>)", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("<body>", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 13 Tf 0.4 0 0.4 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /HelvB 13 Tf 0.4 0 0.4 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("<52696368> Tj", pdf, StringComparison.Ordinal);
        Assert.Contains("<74657874> Tj", pdf, StringComparison.Ordinal);
        Assert.Contains("7365636F6E64", pdf, StringComparison.Ordinal);
        Assert.Contains("6C696E65", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("<526963682074657874> Tj", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenRendersRichContentsCssSpanAppearance() {
        byte[] annotated = BuildFreeTextRichCssContentsPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("font-weight:700", beforePdf, StringComparison.Ordinal);
        Assert.Contains("text-decoration:underline", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("/HelvBI", pdf, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica-BoldOblique", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /HelvBI 11 Tf 0 0.4 0.8 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("<7374796C6564> Tj", pdf, StringComparison.Ordinal);
        Assert.Contains("0 0.4 0.8 RG", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfInspector_ExposesFreeTextDefaultStyleMetadata() {
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(BuildFreeTextDefaultStylePdfWithoutAppearance()).GetAnnotationsBySubtype("FreeText"));

        Assert.True(annotation.HasFreeTextAppearanceMetadata);
        Assert.Null(annotation.DefaultAppearance);
        Assert.Equal("font: Helvetica 15pt; color: #336699; text-align: right", annotation.DefaultStyle);
        Assert.Null(annotation.RichContents);
        Assert.Null(annotation.RichContentsPlainText);
        Assert.Equal(15D, annotation.EffectiveFontSize);
        Assert.NotNull(annotation.EffectiveTextColor);
        Assert.Equal(0.2D, annotation.EffectiveTextColor.Value.R, 3);
        Assert.Equal(0.4D, annotation.EffectiveTextColor.Value.G, 3);
        Assert.Equal(0.6D, annotation.EffectiveTextColor.Value.B, 3);
        Assert.Equal(PdfAlign.Right, annotation.EffectiveTextAlign);
    }

    [Fact]
    public void PdfInspector_ExposesFreeTextRichContentsMetadata() {
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(BuildFreeTextRichContentsPdfWithoutAppearance()).GetAnnotationsBySubtype("FreeText"));

        Assert.True(annotation.HasFreeTextAppearanceMetadata);
        Assert.Equal("font-size: 13pt; color: rgb(102, 0, 102)", annotation.DefaultStyle);
        Assert.Equal("<body><p>Rich <b>text</b><br/>second &amp; line</p></body>", annotation.RichContents);
        Assert.Equal("Rich text\nsecond & line", annotation.RichContentsPlainText);
        Assert.Equal(13D, annotation.EffectiveFontSize);
        Assert.NotNull(annotation.EffectiveTextColor);
        Assert.Equal(0.4D, annotation.EffectiveTextColor.Value.R, 3);
        Assert.Equal(0D, annotation.EffectiveTextColor.Value.G, 3);
        Assert.Equal(0.4D, annotation.EffectiveTextColor.Value.B, 3);
        Assert.Null(annotation.EffectiveTextAlign);
    }

    [Fact]
    public void PdfInspector_ExposesFreeTextBorderAndCalloutMetadata() {
        PdfAnnotation underline = Assert.Single(PdfInspector.Inspect(BuildFreeTextUnderlineAnnotationPdfWithoutAppearance()).GetAnnotationsBySubtype("FreeText"));
        Assert.True(underline.HasVisualStyleMetadata);
        Assert.Equal(2D, underline.BorderWidth);
        Assert.Equal("Underline", underline.BorderStyle);
        Assert.Empty(underline.BorderDashPattern);
        Assert.Equal(new[] { 0.95D, 0.98D, 1D }, underline.InteriorColor);

        PdfAnnotation callout = Assert.Single(PdfInspector.Inspect(BuildFreeTextCalloutAnnotationPdfWithoutAppearance()).GetAnnotationsBySubtype("FreeText"));
        Assert.True(callout.HasVisualStyleMetadata);
        Assert.Equal(2D, callout.BorderWidth);
        Assert.Equal(new[] { 20D, 30D, 60D, 40D, 100D, 60D }, callout.CalloutLine);
        Assert.Equal("ClosedArrow", callout.CalloutLineEnding);
    }

    [Fact]
    public void PdfInspector_ExposesFreeTextCloudyBorderAndRectangleDifferencesMetadata() {
        PdfAnnotation annotation = Assert.Single(PdfInspector.Inspect(BuildFreeTextCloudyBorderAnnotationPdfWithoutAppearance()).GetAnnotationsBySubtype("FreeText"));

        Assert.True(annotation.HasVisualStyleMetadata);
        Assert.Equal(2D, annotation.BorderWidth);
        Assert.Equal("Cloudy", annotation.BorderEffectStyle);
        Assert.Equal(1.5D, annotation.BorderEffectIntensity);
        Assert.Equal(new[] { 12D, 10D, 8D, 6D }, annotation.RectangleDifferences);
        Assert.Equal(new[] { 0.95D, 0.98D, 1D }, annotation.InteriorColor);
    }

    [Fact]
    public void PdfInspector_ExposesTextMarkupAndLinePathGeometryMetadata() {
        PdfAnnotation highlight = Assert.Single(PdfInspector.Inspect(BuildHighlightAnnotationPdfWithQuadPointsWithoutAppearance()).GetAnnotationsBySubtype("Highlight"));
        Assert.True(highlight.HasPathGeometryMetadata);
        Assert.Equal(new[] { 30D, 100D, 90D, 100D, 30D, 92D, 90D, 92D, 100D, 90D, 150D, 90D, 100D, 82D, 150D, 82D }, highlight.QuadPoints);

        PdfAnnotation line = Assert.Single(PdfInspector.Inspect(BuildLineAnnotationPdfWithoutAppearance()).GetAnnotationsBySubtype("Line"));
        Assert.True(line.HasPathGeometryMetadata);
        Assert.Equal(new[] { 40D, 100D, 140D, 100D }, line.LineCoordinates);
        Assert.Equal("OpenArrow", line.LineStartEnding);
        Assert.Equal("ClosedArrow", line.LineEndEnding);
    }

    [Fact]
    public void PdfInspector_ExposesPolygonPolylineAndInkPathGeometryMetadata() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildPathGeometryAnnotationPdfWithoutAppearances());

        PdfAnnotation polygon = Assert.Single(info.GetAnnotationsBySubtype("Polygon"));
        Assert.True(polygon.HasPathGeometryMetadata);
        Assert.Equal(new[] { 40D, 100D, 80D, 130D, 120D, 100D }, polygon.Vertices);

        PdfAnnotation polyLine = Assert.Single(info.GetAnnotationsBySubtype("PolyLine"));
        Assert.True(polyLine.HasPathGeometryMetadata);
        Assert.Equal(new[] { 40D, 60D, 80D, 80D, 120D, 60D }, polyLine.Vertices);

        PdfAnnotation ink = Assert.Single(info.GetAnnotationsBySubtype("Ink"));
        Assert.True(ink.HasPathGeometryMetadata);
        Assert.Equal(2, ink.InkList.Count);
        Assert.Equal(new[] { 30D, 30D, 60D, 45D, 90D, 30D }, ink.InkList[0]);
        Assert.Equal(new[] { 100D, 30D, 130D, 45D, 160D, 30D }, ink.InkList[1]);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenPreservesNonViewableAnnotations() {
        byte[] annotated = BuildMixedVisibilityAnnotationPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Equal(4, PdfInspector.Inspect(annotated).AnnotationCount);
        Assert.Contains("/F 1", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/F 2", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/F 32", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);
        PdfDocumentInfo after = PdfInspector.Inspect(flattened);

        Assert.True(after.HasAnnotations);
        Assert.Equal(3, after.AnnotationCount);
        Assert.Contains("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("Visible free text", pdf, StringComparison.Ordinal);
        Assert.Contains("Invisible free text", pdf, StringComparison.Ordinal);
        Assert.Contains("Hidden free text", pdf, StringComparison.Ordinal);
        Assert.Contains("No view free text", pdf, StringComparison.Ordinal);
        Assert.Contains("/F 1", pdf, StringComparison.Ordinal);
        Assert.Contains("/F 2", pdf, StringComparison.Ordinal);
        Assert.Contains("/F 32", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextUnderlineBorderStyle() {
        byte[] annotated = BuildFreeTextUnderlineAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/BS << /S /U /W 2 >>", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 RG 2 w 0 1 m 150 1 l S", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.2 0.4 0.8 RG 2 w 1 1 148 42 re S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextInsetBorderStyle() {
        byte[] annotated = BuildFreeTextInsetAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/BS << /S /I /W 2 >>", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("0 0 0 RG 2 w 1 1 m 1 43 l 149 43 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.55 0.55 0.55 RG 2 w 1 1 m 149 1 l 149 43 l S", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0 0 0 RG 2 w 1 1 148 42 re S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextOpacity() {
        byte[] annotated = BuildFreeTextOpacityAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/CA 0.42", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnotationOpacityGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnotationOpacityGs << /Type /ExtGState /CA 0.42 /ca 0.42 >>", pdf, StringComparison.Ordinal);
        Assert.Contains("0.95 0.98 1 rg 0 0 150 44 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextCalloutLine() {
        byte[] annotated = BuildFreeTextCalloutAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/CL [20 30 60 40 100 60]", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/LE /ClosedArrow", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 RG 2 w 10 10 m 50 20 l 90 40 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 rg 10 10 m 18.634 8.448 l 16.888 15.433 l h B", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextRectangleDifferences() {
        byte[] annotated = BuildFreeTextRectangleDifferencesAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/RD [12 10 8 6]", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("0.95 0.98 1 rg 12 6 160 64 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 12 6 cm", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 RG 2 w 1 1 158 62 re S", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf 0.1 0.2 0.3 rg 15 55 Td", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.95 0.98 1 rg 0 0 180 80 re f", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.2 0.4 0.8 RG 2 w 1 1 178 78 re S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesFreeTextCloudyBorderEffect() {
        byte[] annotated = BuildFreeTextCloudyBorderAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/BE << /S /C /I 1.5 >>", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/RD [12 10 8 6]", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("0.95 0.98 1 rg 12 6 160 64 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 12 6 cm", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.4 0.8 RG 2 w 5.25 5.25 m", pdf, StringComparison.Ordinal);
        Assert.Contains("9.083 8.083 12.917 8.083 16.75 5.25 c", pdf, StringComparison.Ordinal);
        Assert.Contains("h S", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 12 Tf 0.1 0.2 0.3 rg 15 55 Td", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.2 0.4 0.8 RG 2 w 1 1 158 62 re S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesHighlightQuadPointsWithoutAppearance() {
        byte[] annotated = BuildHighlightAnnotationPdfWithQuadPointsWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/QuadPoints [30 100 90 100 30 92 90 92 100 90 150 90 100 82 150 82]", beforePdf, StringComparison.Ordinal);
        Assert.Equal(1, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/BM /Multiply", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.35", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.8 0.1 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("10 20 m 70 20 l 70 12 l 10 12 l h f", pdf, StringComparison.Ordinal);
        Assert.Contains("80 10 m 130 10 l 130 2 l 80 2 l h f", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0.8 0.1 rg 0 0 140 30 re f", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesHighlightOpacity() {
        byte[] annotated = BuildHighlightAnnotationPdfWithOpacityWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/CA 0.6", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /Highlight", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs gs", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOHighlightGs << /Type /ExtGState /BM /Multiply /CA 0.6 /ca 0.6 >>", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/CA 0.35 /ca 0.35", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.8 0.1 rg", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesUnderlineAndStrikeOutWithoutAppearance() {
        byte[] annotated = BuildTextMarkupLineAnnotationPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Underline", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /StrikeOut", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Underline", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /StrikeOut", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.1 0.3 0.9 RG 1 w", pdf, StringComparison.Ordinal);
        Assert.Contains("10 12 m 70 12 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.9 0.1 0.1 RG 1 w", pdf, StringComparison.Ordinal);
        Assert.Contains("80 6.4 m 130 6.4 l S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesSquigglyWithoutAppearance() {
        byte[] annotated = BuildSquigglyAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Squiggly", beforePdf, StringComparison.Ordinal);
        Assert.Equal(1, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Squiggly", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.6 0.1 RG 1 w", pdf, StringComparison.Ordinal);
        Assert.Contains("10 13.44 m 12.88 14.88 l 15.76 12 l", pdf, StringComparison.Ordinal);
        Assert.Contains("70 13.44 l S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesSquareAndCircleWithoutAppearance() {
        byte[] annotated = BuildShapeAnnotationPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Square", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Circle", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Square", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Circle", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.9 0.9 1 rg 0.3 0.2 0.8 RG 2 w 1 1 98 48 re B", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.2 0.1 RG 4 w [6 2] 0 d 78 30 m", pdf, StringComparison.Ordinal);
        Assert.Contains("c S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesCloudyShapeBorderEffects() {
        byte[] annotated = BuildCloudyShapeAnnotationPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/BE << /S /C /I 1.5 >>", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Square", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Circle", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.9 0.95 1 rg 1 1 98 48 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.3 0.7 RG 2 w 5.25 5.25 m", pdf, StringComparison.Ordinal);
        Assert.Contains("8.979 8.083 12.708 8.083 16.438 5.25 c", pdf, StringComparison.Ordinal);
        Assert.Contains("0.7 0.2 0.1 RG 3 w 73.5 30 m", pdf, StringComparison.Ordinal);
        Assert.Contains("h S", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.9 0.95 1 rg 0.2 0.3 0.7 RG 2 w 1 1 98 48 re B", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesLineWithoutAppearance() {
        byte[] annotated = BuildLineAnnotationPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Line", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/LE [/OpenArrow /ClosedArrow]", beforePdf, StringComparison.Ordinal);
        Assert.Equal(1, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Line", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.1 0.2 0.7 RG 2 w 20 20 m 120 20 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("28 16.4 m 20 20 l 28 23.6 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.1 0.2 0.7 rg 120 20 m 112 23.6 l 112 16.4 l h B", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesStandardLineEndingsWithoutAppearance() {
        byte[] annotated = BuildLineEndingAnnotationsPdfWithoutAppearance();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Equal(4, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Line", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot4 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.4 0.1 rg 20 16 m 28 16 l 28 24 l 20 24 l h B", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.4 0.1 rg 120 20 m 116 24 l 112 20 l 116 16 l h B", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.4 0.1 rg 28 20 m 28 22.209 26.209 24 24 24 c", pdf, StringComparison.Ordinal);
        Assert.Contains("120 23.6 m 120 16.4 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("20 16.4 m 28 20 l 20 23.6 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.4 0.1 rg 112 20 m 120 23.6 l 120 16.4 l h B", pdf, StringComparison.Ordinal);
        Assert.Contains("23.6 23.6 m 16.4 16.4 l S", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void TextAnnotationValidation_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().TextAnnotation(null!));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().TextAnnotation(" "));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().TextAnnotation("note", width: 0));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().TextAnnotation("note", align: PdfAlign.Justify));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().TextAnnotation("note", icon: (PdfTextAnnotationIcon)999));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().FreeTextAnnotation("note", 0, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().FreeTextAnnotation("note", 120, 20, fontSize: 0));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().FreeTextAnnotation("note", 120, 20, align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().FreeTextAnnotation("note", 120, 20, textAlign: PdfAlign.Justify));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().FreeTextAnnotation("note", 120, 20, padding: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().FreeTextAnnotation("note", 120, 20, lineHeight: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().HighlightAnnotation("note", 0, 20));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().HighlightAnnotation("note", 120, 20, align: PdfAlign.Justify));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.TextAnnotation(" ", 10, 10)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas => canvas.TextAnnotation("note", 10, 10, width: 0)));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.FreeTextAnnotation(" ", 10, 10, 120, 20)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas => canvas.FreeTextAnnotation("note", 10, 10, 0, 20)));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.FreeTextAnnotation("note", 10, 10, 120, 20, textAlign: PdfAlign.Justify)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas => canvas.FreeTextAnnotation("note", 10, 10, 120, 20, padding: -1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas => canvas.FreeTextAnnotation("note", 10, 10, 120, 20, lineHeight: 0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create().Canvas(canvas => canvas.HighlightAnnotation("note", 10, 10, 0, 20)));
    }

    private static int CountTextAnnotationOccurrences(string pdf) {
        int count = 0;
        int index = 0;
        while ((index = pdf.IndexOf("/Subtype /Text", index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += "/Subtype /Text".Length;
        }

        return count;
    }

    private static byte[] ConvertFirstNormalAppearanceToStateDictionary(byte[] source) {
        const string prefix = "/AP << /N ";
        const string suffix = " 0 R >>";
        string pdf = Encoding.ASCII.GetString(source);
        int start = pdf.IndexOf(prefix, StringComparison.Ordinal);
        Assert.True(start >= 0);
        int objectNumberStart = start + prefix.Length;
        int end = pdf.IndexOf(suffix, objectNumberStart, StringComparison.Ordinal);
        Assert.True(end > objectNumberStart);

        string appearanceObjectNumber = pdf.Substring(objectNumberStart, end - objectNumberStart);
        string replacement = "/AS /Selected /AP << /N << /Selected " +
            appearanceObjectNumber +
            " 0 R /Off " +
            appearanceObjectNumber +
            " 0 R >> >>";

        string updated = pdf.Substring(0, start) +
            replacement +
            pdf.Substring(end + suffix.Length);
        return Encoding.ASCII.GetBytes(updated);
    }

    private static byte[] BuildVisualAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 160 64] /Contents (Synthetic free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /Q 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Highlight /Rect [20 80 140 94] /Contents (Synthetic highlight) /C [1 0.9 0.1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextDefaultAppearanceColorPdfWithoutAppearance(string defaultAppearance, string contents) {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 180 74] /Contents (" + contents + ") /DA (" + defaultAppearance + ") /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextDefaultStylePdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 84] /Contents (Styled DS text) /DS (font: Helvetica 15pt; color: #336699; text-align: right) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextRichContentsPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 100] /DS (font-size: 13pt; color: rgb(102, 0, 102)) /RC (<body><p>Rich <b>text</b><br/>second &amp; line</p></body>) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextRichCssContentsPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 220 100] /DS (font-size: 13pt; color: rgb(102, 0, 102)) /RC (<body><p>Plain <span style=\"font-weight:700; font-style:italic; text-decoration:underline; color:#0066cc; font-size:11pt\">styled</span> text</p></body>) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMixedVisibilityAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R 8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 160 64] /Contents (Visible free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 80 160 124] /Contents (Invisible free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /F 1 >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 140 160 184] /Contents (Hidden free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /F 2 >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 200 160 244] /Contents (No view free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /F 32 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextUnderlineAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 160 64] /Contents (Underline free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /BS << /S /U /W 2 >> /C [0.2 0.4 0.8] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextInsetAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 160 64] /Contents (Inset free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /BS << /S /I /W 2 >> /C [0 0 0] /IC [0.95 0.98 1] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextOpacityAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 160 64] /Contents (Opacity free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /CA 0.42 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextCalloutAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 100] /Contents (Callout free text) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 2] /C [0.2 0.4 0.8] /CL [20 30 60 40 100 60] /LE /ClosedArrow >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextRectangleDifferencesAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 100] /Contents (Inset text box) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 2] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /RD [12 10 8 6] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextCloudyBorderAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 100] /Contents (Cloudy text box) /DA (/Helv 12 Tf 0.1 0.2 0.3 rg) /Border [0 0 2] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /RD [12 10 8 6] /BE << /S /C /I 1.5 >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildHighlightAnnotationPdfWithQuadPointsWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Highlight /Rect [20 80 160 110] /Contents (Two line synthetic highlight) /C [1 0.8 0.1] /QuadPoints [30 100 90 100 30 92 90 92 100 90 150 90 100 82 150 82] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildHighlightAnnotationPdfWithOpacityWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Highlight /Rect [20 80 160 110] /Contents (Opacity highlight) /C [1 0.8 0.1] /CA 0.6 /QuadPoints [30 100 90 100 30 92 90 92] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTextMarkupLineAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Underline /Rect [20 80 160 110] /Contents (Synthetic underline) /C [0.1 0.3 0.9] /QuadPoints [30 100 90 100 30 92 90 92] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /StrikeOut /Rect [20 80 160 110] /Contents (Synthetic strikeout) /C [0.9 0.1 0.1] /QuadPoints [100 90 150 90 100 82 150 82] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSquigglyAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Squiggly /Rect [20 80 160 110] /Contents (Synthetic squiggly) /C [0.2 0.6 0.1] /QuadPoints [30 100 90 100 30 92 90 92] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildShapeAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Square /Rect [20 80 120 130] /Contents (Synthetic square) /C [0.3 0.2 0.8] /IC [0.9 0.9 1] /Border [0 0 2] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Circle /Rect [140 80 220 140] /Contents (Synthetic circle) /C [0.8 0.2 0.1] /Border [0 0 3] /BS << /S /D /W 4 /D [6 2] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCloudyShapeAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Square /Rect [20 80 120 130] /Contents (Cloudy square) /C [0.2 0.3 0.7] /IC [0.9 0.95 1] /Border [0 0 2] /BE << /S /C /I 1.5 >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Circle /Rect [140 80 220 140] /Contents (Cloudy circle) /C [0.7 0.2 0.1] /Border [0 0 3] /BE << /S /C /I 2 >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildLineAnnotationPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 80 160 120] /Contents (Synthetic line) /L [40 100 140 100] /C [0.1 0.2 0.7] /Border [0 0 2] /LE [/OpenArrow /ClosedArrow] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildLineEndingAnnotationsPdfWithoutAppearance() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R 8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 80 160 120] /Contents (Line endings 1) /L [40 100 140 100] /C [0.1 0.2 0.7] /IC [0.8 0.4 0.1] /Border [0 0 2] /LE [/Square /Diamond] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 130 160 170] /Contents (Line endings 2) /L [40 150 140 150] /C [0.1 0.2 0.7] /IC [0.8 0.4 0.1] /Border [0 0 2] /LE [/Circle /Butt] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 180 160 220] /Contents (Line endings 3) /L [40 200 140 200] /C [0.1 0.2 0.7] /IC [0.8 0.4 0.1] /Border [0 0 2] /LE [/ROpenArrow /RClosedArrow] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 230 160 270] /Contents (Line endings 4) /L [40 250 140 250] /C [0.1 0.2 0.7] /IC [0.8 0.4 0.1] /Border [0 0 2] /LE [/Slash /None] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPathGeometryAnnotationPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Polygon /Rect [20 90 140 140] /Contents (Polygon geometry) /Vertices [40 100 80 130 120 100] /C [0.2 0.3 0.8] /IC [0.9 0.9 1] /Border [0 0 2] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /PolyLine /Rect [20 50 140 90] /Contents (Polyline geometry) /Vertices [40 60 80 80 120 60] /C [0.8 0.3 0.2] /Border [0 0 2] /LE [/None /OpenArrow] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Ink /Rect [20 20 180 60] /Contents (Ink geometry) /InkList [[30 30 60 45 90 30] [100 30 130 45 160 30]] /C [0.1 0.2 0.7] /Border [0 0 2] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
