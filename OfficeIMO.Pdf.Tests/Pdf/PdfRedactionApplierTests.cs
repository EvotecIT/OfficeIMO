using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionApplierTests {
    [Fact]
    public void Apply_RemovesMatchedTextAndKeepsUnmatchedTextExtractable() {
        byte[] source = BuildRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");

        PdfRedactionPlan plan = PdfRedactionPlanner.Plan(source, new[] { area });
        Assert.True(plan.HasMatches);
        Assert.Contains(plan.Matches, match => match.Text != null && match.Text.Contains("Secret account", StringComparison.Ordinal));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Secret account", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);

        PdfRedactionPlan redactedPlan = PdfRedactionPlanner.Plan(redacted, new[] { area });
        Assert.DoesNotContain(redactedPlan.Matches, match => match.Text != null && match.Text.Contains("Secret account", StringComparison.Ordinal));
    }

    [Fact]
    public void ApplyRedactions_FacadeReturnsRedactedDocumentAndTryResult() {
        byte[] source = BuildRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        PdfDocument document = PdfDocument.Open(source);

        PdfDocument redacted = document.ApplyRedactions(new[] { area });
        PdfOperationResult<PdfDocument> result = document.TryApplyRedactions(new[] { area });

        Assert.DoesNotContain("Secret account", redacted.Read.Text(), StringComparison.Ordinal);
        Assert.True(result.Succeeded);
        Assert.DoesNotContain("Secret account", result.RequireValue().Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_DecodesOctalEscapesBeforeMatchingTextObjects() {
        byte[] source = BuildOctalRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_DecodesNestedLiteralStringsBeforeMatchingTextObjects() {
        byte[] source = BuildNestedLiteralRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Account (secret)");
        Assert.Contains("Account (secret)", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Account (secret)", text, StringComparison.Ordinal);
        Assert.DoesNotContain("secret", text, StringComparison.Ordinal);
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByOuterGraphicsTransform() {
        byte[] source = BuildTransformedTextRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Transformed secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Transformed secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Transformed secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByTransformFromPriorContentStream() {
        byte[] source = BuildSplitTransformedTextRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split transformed secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split transformed secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split transformed secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsTextPositionedByTransformOperandsSplitAcrossContentStreams() {
        byte[] source = BuildSplitTransformOperandRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split operand secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split operand secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split operand secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_IgnoresEndTextOperatorInsideLiteralStringsWhenScrubbingTextObjects() {
        byte[] source = BuildLiteralEndTextOperatorRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "SSN ET 123");
        Assert.Contains("SSN ET 123", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("SSN ET 123", text, StringComparison.Ordinal);
        Assert.DoesNotContain("SSN ET 123", raw, StringComparison.Ordinal);
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_UsesFontDecoderForLiteralRedactionText() {
        byte[] source = BuildToUnicodeLiteralRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", ExtractLogicalText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = ExtractLogicalText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsMatchedTextInsideFormXObjects() {
        byte[] source = BuildFormXObjectRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Secret account 123-45");
        Assert.Contains("Secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.DoesNotContain("Secret account", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PreservesTokensSplitAcrossPageContentStreamsWhenLocatingForms() {
        byte[] source = BuildSplitFormTransformOperandRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Split form secret");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        Assert.DoesNotContain("Split form secret", PdfTextExtractor.ExtractAllText(redacted), StringComparison.Ordinal);
        Assert.DoesNotContain("Split form secret", PdfEncoding.Latin1GetString(redacted), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScrubsMatchedTextInsideNestedFormXObjects() {
        byte[] source = BuildNestedFormXObjectRedactionSource();
        PdfRedactionArea area = FindAreaForText(source, "Nested secret account 123-45");
        Assert.Contains("Nested secret account 123-45", PdfTextExtractor.ExtractAllText(source), StringComparison.Ordinal);

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("Nested secret", text, StringComparison.Ordinal);
        Assert.DoesNotContain("123-45", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Nested secret account 123-45", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ClonesSharedPageContentBeforeScrubbingMatchedText() {
        byte[] source = BuildSharedPageContentPdf();
        PdfRedactionArea area = FindAreasForText(source, "Shared page secret").Single(redaction => redaction.PageNumber == 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Shared page secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Shared page secret"));
    }

    [Fact]
    public void ReplacePageContentReferenceAtIndex_ReplacesOnlyTheSelectedRepeatedOccurrence() {
        var first = new PdfReference(5, 0);
        var middle = new PdfReference(6, 0);
        var repeated = new PdfReference(5, 0);
        var replacement = new PdfReference(7, 0);
        var contents = new PdfArray();
        contents.Items.Add(first);
        contents.Items.Add(middle);
        contents.Items.Add(repeated);
        var page = new PdfDictionary();
        page.Items["Contents"] = contents;

        PdfRedactionApplier.ReplacePageContentReferenceAtIndex(
            new Dictionary<int, PdfIndirectObject>(),
            page,
            contents,
            contentIndex: 2,
            replacement);

        Assert.Same(first, contents.Items[0]);
        Assert.Same(middle, contents.Items[1]);
        Assert.Same(replacement, contents.Items[2]);
    }

    [Fact]
    public void Apply_ClonesSharedFormXObjectBeforeScrubbingMatchedText() {
        byte[] source = BuildSharedFormXObjectTextPdf();
        PdfRedactionArea area = FindAreasForText(source, "Shared form secret").Single(redaction => redaction.PageNumber == 1);
        Assert.Equal(2, CountOccurrences(PdfTextExtractor.ExtractAllText(source), "Shared form secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Shared form secret"));
    }

    [Fact]
    public void Apply_IsolatesExistingContentBeforePaintingRedactionOverlay() {
        byte[] source = BuildLeakingGraphicsStateRedactionSource();
        var area = new PdfRedactionArea(1, 40, 40, 80, 24, "manual");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("\nq\n", raw, StringComparison.Ordinal);
        Assert.Contains("0 0 1 1 re W n", raw, StringComparison.Ordinal);
        Assert.Contains("\nQ\n", raw, StringComparison.Ordinal);
        Assert.Contains("40 40 80 24 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_IsolatesContentArrayAsSingleSequenceBeforeOverlay() {
        byte[] source = BuildSplitContentStateRedactionSource();
        var area = new PdfRedactionArea(1, 40, 40, 80, 24, "manual");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("\nq\n", raw, StringComparison.Ordinal);
        Assert.Contains("\nQ\n", raw, StringComparison.Ordinal);
        Assert.Contains("/F1 12 Tf", raw, StringComparison.Ordinal);
        Assert.Contains("(Visible split text) Tj", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_ScopesDuplicateTextRedactionToIntersectingInstance() {
        byte[] source = BuildDuplicateRedactionSource();
        PdfRedactionArea area = FindAreaForTextOccurrence(source, "Repeat secret", occurrenceFromTop: 1);
        string originalText = PdfTextExtractor.ExtractAllText(source);
        Assert.Equal(2, CountOccurrences(originalText, "Repeat secret"));

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string text = PdfTextExtractor.ExtractAllText(redacted);

        Assert.Equal(1, CountOccurrences(text, "Repeat secret"));
        Assert.Contains("Visible before", text, StringComparison.Ordinal);
        Assert.Contains("Visible after", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_PrunesRemovedAnnotationAppearanceStreams() {
        byte[] source = BuildAnnotationAppearanceRedactionSource();
        var area = new PdfRedactionArea(1, 20, 20, 40, 40, "annotation");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.Contains("Sensitive annotation", PdfEncoding.Latin1GetString(source), StringComparison.Ordinal);
        Assert.DoesNotContain("Sensitive annotation", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("Old sensitive appearance", raw, StringComparison.Ordinal);
        Assert.Empty(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("FreeText"));
    }

    [Fact]
    public void Apply_ClearsParentPopupReferenceWhenRedactingPopupAnnotation() {
        byte[] source = BuildIndirectAnnotationWithPopupPdf();
        var area = new PdfRedactionArea(1, 100, 100, 60, 60, "popup");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });
        string raw = PdfEncoding.Latin1GetString(redacted);

        Assert.DoesNotContain("/Popup", raw, StringComparison.Ordinal);
        Assert.Single(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("Text"));
        Assert.Empty(PdfInspector.Inspect(redacted).GetAnnotationsBySubtype("Popup"));
    }

    private static byte[] BuildRedactionSource() {
        return PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Visible before"))
            .Paragraph(paragraph => paragraph.Text("Secret account 123-45"))
            .Paragraph(paragraph => paragraph.Text("Visible after"))
            .ToBytes();
    }

    private static byte[] BuildOctalRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "0 -18 Td",
            "(Secret\\040account\\040123-45) Tj",
            "0 -18 Td",
            "(Visible after) Tj",
            "ET"
        });
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(streamContent).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNestedLiteralRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 702 Td",
            "(Account (secret)) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 650 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildTransformedTextRedactionSource() {
        const string streamContent = "q\n2 0 0 2 100 100 cm\nBT\n/F1 12 Tf\n0 0 Td\n(Transformed secret) Tj\nET\nQ\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitTransformedTextRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n2 0 0 2 100 100 cm\n")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("BT\n/F1 12 Tf\n0 0 Td\n(Split transformed secret) Tj\nET\nQ\n"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitTransformOperandRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n2 0 0 2 1")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("00 100 cm\nBT\n/F1 12 Tf\n0 0 Td\n(Split operand secret) Tj\nET\nQ\n"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildLiteralEndTextOperatorRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 702 Td",
            "(SSN ET 123) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 650 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildToUnicodeLiteralRedactionSource() {
        string secret = "Secret account 123-45";
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible before) Tj",
            "0 -18 Td",
            "(" + EncodeLiteralGlyphBytes(secret) + ") Tj",
            "0 -18 Td",
            "(Visible after) Tj",
            "ET"
        });
        string cmap = BuildSingleByteToUnicodeCMap(secret);
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /AAAAAA+Helvetica /Encoding /WinAnsiEncoding /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildFormXObjectRedactionSource() {
        string pageContent = "q\n1 0 0 1 72 700 cm\n/Fm1 Do\nQ\n";
        string formContent = "BT\n/F1 12 Tf\n0 0 Td\n(Secret account 123-45) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitFormTransformOperandRedactionSource() {
        const string formContent = "BT\n/F1 12 Tf\n0 0 Td\n(Split form secret) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Fm1 7 0 R >> >> /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("q\n1 0 0 1 1")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("00 100 cm\n/Fm1 Do\nQ\n")),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(formContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildNestedFormXObjectRedactionSource() {
        string pageContent = "q\n1 0 0 1 72 700 cm\n/FmOuter Do\nQ\n";
        string outerFormContent = "q\n1 0 0 1 0 0 cm\n/FmInner Do\nQ\n";
        string innerFormContent = "BT\n/F1 12 Tf\n0 0 Td\n(Nested secret account 123-45) Tj\nET\n";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /FmOuter 6 0 R >> >> /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(outerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /XObject << /FmInner 7 0 R >> >>"),
            BuildStreamObject(7, Encoding.ASCII.GetBytes(innerFormContent), "/Type /XObject /Subtype /Form /BBox [0 0 220 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildLeakingGraphicsStateRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "q",
            "0 0 1 1 re W n",
            "BT",
            "/F1 12 Tf",
            "72 720 Td",
            "(Visible page text) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildSplitContentStateRedactionSource() {
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents [5 0 R 6 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes("/F1 12 Tf")),
            BuildStreamObject(6, Encoding.ASCII.GetBytes("72 720 Td (Visible split text) Tj"))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildDuplicateRedactionSource() {
        string streamContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "72 740 Td",
            "(Visible before) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 700 Td",
            "(Repeat secret) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 660 Td",
            "(Repeat secret) Tj",
            "ET",
            "BT",
            "/F1 12 Tf",
            "72 620 Td",
            "(Visible after) Tj",
            "ET"
        });
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(streamContent))
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildAnnotationAppearanceRedactionSource() {
        string pageContent = "BT\n/F1 12 Tf\n72 720 Td\n(Visible page text) Tj\nET";
        string appearanceContent = "BT /F1 12 Tf 0 0 Td (Old sensitive appearance Sensitive annotation) Tj ET";
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /Annots [6 0 R] /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(pageContent)),
            "6 0 obj\n<< /Type /Annot /Subtype /FreeText /Rect [20 20 60 60] /Contents (Sensitive annotation) /AP << /N 7 0 R >> >>\nendobj",
            BuildStreamObject(7, Encoding.ASCII.GetBytes(appearanceContent), "/Type /XObject /Subtype /Form /BBox [0 0 40 40] /Resources << /Font << /F1 4 0 R >> >>")
        };

        return BuildPdf(objects, rootObjectNumber: 1);
    }

    private static byte[] BuildFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildInheritedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /Resources << /XObject << /Fm1 5 0 R >> >> >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Inherited form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildSharedPageContentPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n72 120 Td\n(Shared page secret) Tj\nET")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildSharedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Fm1 6 0 R >> >> /Contents 8 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Shared form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 5 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/Fm1 Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildAliasedFormXObjectTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /FmUnused 5 0 R /FmPainted 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 12 Tf\n0 0 Td\n(Aliased form secret) Tj\nET", "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Resources << /Font << /F1 4 0 R >> >>"),
            BuildStream("q\n1 0 0 1 100 100 cm\n/FmPainted Do\nQ")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildLargeTextPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            BuildStream("BT\n/F1 48 Tf\n72 100 Td\n(Large secret heading) Tj\nET")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildDirectAnnotationWithPopupPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Direct redaction note) /Popup 5 0 R >> 5 0 R] /Contents 4 0 R >>",
            BuildStream("BT\n/F1 12 Tf\n72 720 Td\n(Annotation carrier) Tj\nET"),
            "<< /Type /Annot /Subtype /Popup /Rect [45 20 120 80] >>"
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static byte[] BuildIndirectAnnotationWithPopupPdf() {
        var objects = new List<string> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [4 0 R 5 0 R] /Contents 6 0 R >>",
            "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (Keep parent note) /Popup 5 0 R >>",
            "<< /Type /Annot /Subtype /Popup /Rect [100 100 160 160] /Parent 4 0 R >>",
            BuildStream("BT\n/F1 12 Tf\n72 720 Td\n(Annotation carrier) Tj\nET")
        };

        return Encoding.ASCII.GetBytes(BuildPdf(objects));
    }

    private static string BuildStream(string content, string dictionaryEntries = "") {
        byte[] bytes = Encoding.ASCII.GetBytes(content);
        return "<< " + dictionaryEntries + (dictionaryEntries.Length == 0 ? string.Empty : " ") + "/Length " + bytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";
    }

    private static string BuildPdf(IReadOnlyList<string> objects) {
        var builder = new StringBuilder();
        builder.AppendLine("%PDF-1.7");
        for (int i = 0; i < objects.Count; i++) {
            builder.Append((i + 1).ToString(CultureInfo.InvariantCulture)).AppendLine(" 0 obj");
            builder.AppendLine(objects[i]);
            builder.AppendLine("endobj");
        }

        builder.AppendLine("trailer");
        builder.Append("<< /Root 1 0 R /Size ").Append(objects.Count + 1).AppendLine(" >>");
        builder.AppendLine("startxref");
        builder.AppendLine("123");
        builder.AppendLine("%%EOF");
        return builder.ToString();
    }

    private static PdfRedactionArea FindAreaForText(byte[] pdf, string text) {
        return FindAreasForText(pdf, text).Single();
    }

    private static PdfRedactionArea[] FindAreasForText(byte[] pdf, string text) {
        return PdfLogicalDocument.Load(pdf)
            .TextBlocks
            .Where(item => item.Text.Contains(text, StringComparison.Ordinal))
            .Select(static block => {
                double x = Math.Min(block.XStart, block.XEnd) - 2D;
                double width = Math.Abs(block.XEnd - block.XStart) + 4D;
                return new PdfRedactionArea(block.PageNumber, x, block.BaselineY - 14D, width, 20D, "secret");
            })
            .ToArray();
    }

    private static int CountOccurrences(string value, string search) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += search.Length;
        }

        return count;
    }

    private static PdfRedactionArea FindAreaForTextOccurrence(byte[] pdf, string text, int occurrenceFromTop) {
        PdfLogicalTextBlock block = PdfLogicalDocument.Load(pdf)
            .TextBlocks
            .Where(item => item.Text.Contains(text, StringComparison.Ordinal))
            .OrderByDescending(item => item.BaselineY)
            .ElementAt(occurrenceFromTop);

        double x = Math.Min(block.XStart, block.XEnd) - 2D;
        double width = Math.Abs(block.XEnd - block.XStart) + 4D;
        return new PdfRedactionArea(block.PageNumber, x, block.BaselineY - 14D, width, 20D, "secret");
    }

    private static string BuildSingleByteToUnicodeCMap(string text) {
        var builder = new StringBuilder();
        builder.Append("/CIDInit /ProcSet findresource begin\n");
        builder.Append("12 dict begin\n");
        builder.Append("begincmap\n");
        builder.Append(text.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
        for (int i = 0; i < text.Length; i++) {
            builder.Append('<')
                .Append((i + 1).ToString("X2", System.Globalization.CultureInfo.InvariantCulture))
                .Append("> <")
                .Append(((int)text[i]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture))
                .Append(">\n");
        }

        builder.Append("endbfchar\n");
        builder.Append("endcmap\n");
        builder.Append("CMapName currentdict /CMap defineresource pop\n");
        builder.Append("end\n");
        builder.Append("end\n");
        return builder.ToString();
    }

    private static string EncodeLiteralGlyphBytes(string text) {
        var builder = new StringBuilder(text.Length * 4);
        for (int i = 0; i < text.Length; i++) {
            builder.Append('\\')
                .Append(Convert.ToString(i + 1, 8).PadLeft(3, '0'));
        }

        return builder.ToString();
    }

    private static string BuildStreamObject(int objectNumber, byte[] streamBytes, string extraDictionary = "") {
        string dictionary = "<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(extraDictionary)) {
            dictionary += " " + extraDictionary.Trim();
        }

        dictionary += " >>";
        return objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" +
            dictionary + "\nstream\n" +
            Encoding.ASCII.GetString(streamBytes) + "\nendstream\nendobj";
    }

    private static byte[] BuildPdf(IReadOnlyList<string> objects, int rootObjectNumber) {
        var offsets = new Dictionary<int, int>();
        using var stream = new MemoryStream();
        using var writer = new StreamWriter(stream, Encoding.ASCII, 1024, leaveOpen: true);

        writer.WriteLine("%PDF-1.4");
        writer.Flush();
        int maxObjectNumber = 0;
        foreach (string obj in objects) {
            int objectNumber = ReadObjectNumber(obj);
            offsets[objectNumber] = (int)stream.Position;
            maxObjectNumber = Math.Max(maxObjectNumber, objectNumber);
            writer.WriteLine(obj);
            writer.Flush();
        }

        int xrefOffset = (int)stream.Position;
        writer.WriteLine("xref");
        writer.WriteLine("0 " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("0000000000 65535 f ");
        for (int i = 1; i <= maxObjectNumber; i++) {
            if (offsets.TryGetValue(i, out int offset)) {
                writer.WriteLine(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n ");
            } else {
                writer.WriteLine("0000000000 65535 f ");
            }
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + (maxObjectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
        return stream.ToArray();
    }

    private static int ReadObjectNumber(string obj) {
        int space = obj.IndexOf(' ');
        return int.Parse(obj.Substring(0, space), System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string ExtractLogicalText(byte[] pdf) {
        return string.Join(
            Environment.NewLine,
            PdfLogicalDocument.Load(pdf).TextBlocks.Select(item => item.Text));
    }
}
