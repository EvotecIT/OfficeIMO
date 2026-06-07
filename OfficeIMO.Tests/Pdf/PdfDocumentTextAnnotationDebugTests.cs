using System;
using System.Text;
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
        Assert.Contains("/Contents (Review flow anchor)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Canvas note)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Flow free text wraps across several words for reviewers)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Canvas free text wraps as a right aligned reviewer callout)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Flow highlight)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Canvas highlight)", pdf, StringComparison.Ordinal);
        Assert.Contains("/C [1 0 0]", pdf, StringComparison.Ordinal);
        Assert.Contains("/C [0.2 0.4 0.8]", pdf, StringComparison.Ordinal);
        Assert.Contains("/IC [0.95 0.98 1]", pdf, StringComparison.Ordinal);
        Assert.Contains("/QuadPoints [", pdf, StringComparison.Ordinal);
        Assert.Equal(4, CountOccurrences(pdf, "/AP << /N "));
        Assert.True(CountOccurrences(pdf, "BT /Helv") >= 4);
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
        Assert.Contains("BT /Helv", pdf, StringComparison.Ordinal);
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
        Assert.Contains("BT /Helv", pdf, StringComparison.Ordinal);
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
        Assert.Contains("BT /Helv", pdf, StringComparison.Ordinal);
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
        Assert.Contains("1 0.9 0.1 rg 0 0 120 14 re f", pdf, StringComparison.Ordinal);
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
        Assert.Contains("1 0.8 0.1 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("10 20 m 70 20 l 70 12 l 10 12 l h f", pdf, StringComparison.Ordinal);
        Assert.Contains("80 10 m 130 10 l 130 2 l 80 2 l h f", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0.8 0.1 rg 0 0 140 30 re f", pdf, StringComparison.Ordinal);
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
        Assert.Contains("0.8 0.2 0.1 RG 3 w 78.5 30 m", pdf, StringComparison.Ordinal);
        Assert.Contains("c S", pdf, StringComparison.Ordinal);
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
            "<< /Type /Annot /Subtype /Circle /Rect [140 80 220 140] /Contents (Synthetic circle) /C [0.8 0.2 0.1] /Border [0 0 3] >>",
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
}
