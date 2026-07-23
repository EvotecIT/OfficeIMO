using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfResourceBudgetSecurityTests {
    [Fact]
    public void ToUnicodeCMap_IgnoresSourceCodesLongerThanPdfLimit() {
        const string source = "1 beginbfchar\n<0102030405> <0041>\nendbfchar";

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));

        Assert.Equal("\u0001\u0002\u0003\u0004\u0005", cmap!.MapBytes(new byte[] { 1, 2, 3, 4, 5 }));
    }

    [Fact]
    public void ToUnicodeCMap_RejectsOversizedDestinationAndOutputExpansion() {
        string oversizedDestination = string.Concat(Enumerable.Repeat("0041", 4097));
        string source = "2 beginbfchar\n<01> <" + oversizedDestination + ">\n<02> <0041004200430044>\nendbfchar";

        Assert.True(ToUnicodeCMap.TryParse(Encoding.ASCII.GetBytes(source), out ToUnicodeCMap? cmap));
        Assert.Equal("\u0001", cmap!.MapBytes(new byte[] { 1 }));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            cmap.MapBytes(new byte[] { 2, 2, 2 }, maxOutputCharacters: 10));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void TextContentParser_BoundsRepeatedActualTextExpansion() {
        const string content =
            "/Span /Shared BDC BT /F1 12 Tf (A) Tj ET EMC " +
            "/Span /Shared BDC BT /F1 12 Tf (B) Tj ET EMC";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content,
                static (_, bytes) => Encoding.ASCII.GetString(bytes),
                static (_, bytes) => bytes.Length * 500D,
                actualTextForProperty: static _ => Encoding.ASCII.GetBytes("123456"),
                maxActualTextCharacters: 10));

        Assert.Equal(PdfReadLimitKind.ActualTextCharacters, exception.Kind);
    }

    [Fact]
    public void TextContentParser_BoundsInlineActualTextBeforeDecoding() {
        const string oversized = "/Span << /ActualText (AB) >> BDC BT /F1 12 Tf (X) Tj ET EMC";
        const string boundedUtf16 = "/Span << /ActualText <FEFF00410042> >> BDC BT /F1 12 Tf (X) Tj ET EMC";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                oversized,
                static (_, bytes) => Encoding.ASCII.GetString(bytes),
                static (_, bytes) => bytes.Length * 500D,
                maxActualTextCharacters: 1));
        IReadOnlyList<PdfTextSpan> spans = TextContentParser.Parse(
            boundedUtf16,
            static (_, bytes) => Encoding.ASCII.GetString(bytes),
            static (_, bytes) => bytes.Length * 500D,
            maxActualTextCharacters: 2);

        Assert.Equal(PdfReadLimitKind.ActualTextCharacters, exception.Kind);
        Assert.Equal("AB", Assert.Single(spans).Text);
    }

    [Fact]
    public void TextContentParser_ChargesDiscardedArtifactTextToDecodedBudget() {
        const string content = "/Artifact BMC BT /F1 12 Tf (AB) Tj ET EMC";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content,
                static (_, bytes) => Encoding.ASCII.GetString(bytes),
                static (_, bytes) => bytes.Length * 500D,
                maxDecodedTextCharacters: 1));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void TextContentParser_ChargesFontDecodedTextReplacedByActualText() {
        const string content = "/Span << /ActualText (X) >> BDC BT /F1 12 Tf (AB) Tj ET EMC";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content,
                static (_, bytes) => Encoding.ASCII.GetString(bytes),
                static (_, bytes) => bytes.Length * 500D,
                maxDecodedTextCharacters: 1));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void TextContentParser_ChargesDecodedGlyphsDroppedByThinSpacingFilter() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                "BT /F1 12 Tf (AB) Tj ET",
                static (_, bytes) => bytes.Length == 1 && bytes[0] == (byte)'B'
                    ? "BBBBBB"
                    : Encoding.ASCII.GetString(bytes),
                static (_, bytes) => bytes[0] == (byte)'B' ? 0D : bytes.Length * 500D,
                maxDecodedTextCharacters: 6));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void TextContentParser_BoundsTextRunBufferBeforeDecoding() {
        var budget = new TextContentParser.TextOutputBudget(
            maxActualTextCharacters: 1,
            maxDecodedTextCharacters: 1);

        Assert.Equal(1, budget.GetDecodedTextBufferCapacity(int.MaxValue));
        budget.ChargeDecodedText(1);
        Assert.Equal(0, budget.GetDecodedTextBufferCapacity(int.MaxValue));
    }

    [Fact]
    public void TextContentParser_PassesRemainingBudgetToEachTextRunDecoder() {
        var decoderLimits = new List<int>();
        const string content = "BT /F1 12 Tf (123456789) Tj (abcdefghij) Tj ET";

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            TextContentParser.Parse(
                content,
                static (_, _) => throw new InvalidOperationException("The bounded decoder must be used."),
                static (_, bytes) => bytes.Length * 500D,
                maxDecodedTextCharacters: 10,
                decodeWithFontWithinLimit: (_, bytes, maximumCharacters) => {
                    decoderLimits.Add(maximumCharacters);
                    return PdfWinAnsiEncoding.Decode(bytes, maximumCharacters);
                }));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
        Assert.Contains(1, decoderLimits);
    }

    [Fact]
    public void PdfReadPage_BoundsAggregateContentStreamBytes() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents [4 0 R 5 0 R] >>",
            "<< /Length 5 >>\nstream\nBT ET\nendstream",
            "<< /Length 5 >>\nstream\nBT ET\nendstream");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPageContentBytes = 9 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.PageContentBytes, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_ChargesRepeatedContentStreamReferencesPerOccurrence() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents [4 0 R 4 0 R] >>",
            BuildStream("BT ET"));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPageContentBytes = 9 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.PageContentBytes, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_BoundsNestedFormContentAgainstPageBudget() {
        const string pageContent = "/Fx Do";
        const string formContent = "BT (Nested) Tj ET";
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /XObject << /Fx 5 0 R >> >> /Contents 4 0 R >>",
            BuildStream(pageContent),
            BuildStream(formContent, "/Type /XObject /Subtype /Form /BBox [0 0 100 100]"));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPageContentBytes = pageContent.Length + formContent.Length - 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.PageContentBytes, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_ToDrawingChargesAnnotationAppearancesToPageBudget() {
        const string appearanceContent = "BT (Annotation) Tj ET";
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Annots [5 0 R] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Type /Annot /Subtype /FreeText /Rect [0 0 10 10] /AP << /N 6 0 R >> >>",
            BuildStream(appearanceContent, "/Type /XObject /Subtype /Form /BBox [0 0 10 10]"));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxPageContentBytes = appearanceContent.Length - 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.PageContentBytes, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_SharesActualTextBudgetAcrossRepeatedForms() {
        const string pageContent = "/Fx Do /Fx Do";
        const string formContent = "/Span << /ActualText (123456) >> BDC BT (A) Tj ET EMC";
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /XObject << /Fx 5 0 R >> >> /Contents 4 0 R >>",
            BuildStream(pageContent),
            BuildStream(formContent, "/Type /XObject /Subtype /Form /BBox [0 0 100 100]"));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxActualTextCharacters = 10 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.ActualTextCharacters, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_BoundsNamedActualTextFromRawBytes() {
        const string pageContent = "/Span /MC0 BDC BT /F1 12 Tf (X) Tj ET EMC";
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Properties << /MC0 << /ActualText <FEFF00410042> >> >> >> /Contents 4 0 R >>",
            BuildStream(pageContent));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxActualTextCharacters = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.ActualTextCharacters, exception.Kind);
    }

    [Fact]
    public void PdfReadPage_ToDrawingSharesDecodedTextBudgetAcrossAnnotationAppearances() {
        const string appearanceContent = "BT /F1 12 Tf (123456) Tj ET";
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 8 0 R >> >> /Annots [5 0 R 6 0 R] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Type /Annot /Subtype /FreeText /Rect [0 0 10 10] /Contents (First) /AP << /N 7 0 R >> >>",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 0 30 10] /Contents (Second) /AP << /N 7 0 R >> >>",
            BuildStream(appearanceContent, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /Font << /F1 8 0 R >> >>"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 10 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(pdf, options).Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void StreamDecoder_RejectsOversizedPredictorRowBeforeAllocation() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var decodeParameters = new PdfDictionary();
        decodeParameters.Items["Predictor"] = new PdfNumber(12);
        decodeParameters.Items["Columns"] = new PdfNumber(100_000_000);
        dictionary.Items["DecodeParms"] = decodeParameters;

        byte[] encoded;
        using (var buffer = new MemoryStream()) {
            using (var compressor = new DeflateStream(buffer, CompressionLevel.Optimal, leaveOpen: true)) {
                compressor.WriteByte(0);
            }
            encoded = buffer.ToArray();
        }

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            StreamDecoder.Decode(dictionary, encoded, maxOutputBytes: 64));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(64, exception.Limit);
    }

    [Fact]
    public void PdfReadDocument_BoundsNamedDestinationTreeDepth() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 5 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Kids [6 0 R] >>",
            "<< /Kids [7 0 R] >>",
            "<< /Names [(Target) [3 0 R /Fit]] >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeDepth = 1 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf, options));

        Assert.Equal(PdfReadLimitKind.NameTreeDepth, exception.Kind);
    }

    [Fact]
    public void PdfReadDocument_CountsReferencedNameTreeNodeOnce() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 5 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Target) [3 0 R /Fit]] >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        Assert.Equal("Target", Assert.Single(document.NamedDestinations).Name);
    }

    [Fact]
    public void PdfReadDocument_CountsOnlyUniqueIndirectNameTreeNodes() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Kids [5 0 R 5 0 R] >> >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Target) [3 0 R /Fit]] >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        Assert.Equal("Target", Assert.Single(document.NamedDestinations).Name);
    }

    [Fact]
    public void PdfReadDocument_CountsReferencedCatalogActionNameTreeNodeOnce() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript 5 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Open) 6 0 R] >>",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfCatalogAction action = Assert.Single(document.CatalogActions);
        Assert.Equal("Open", action.Name);
        Assert.Equal("JavaScript", action.ActionType);
    }

    [Fact]
    public void PdfReadDocument_CountsOnlyUniqueIndirectCatalogActionNameTreeNodes() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Kids [5 0 R 5 0 R] >> >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Open) 6 0 R] >>",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        PdfCatalogAction action = Assert.Single(document.CatalogActions);
        Assert.Equal("Open", action.Name);
        Assert.Equal("JavaScript", action.ActionType);
    }

    [Fact]
    public void PdfReadDocument_InvalidGenerationDoesNotHideLaterValidCatalogActionNode() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Kids [5 1 R 5 0 R] >> >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Open) 6 0 R] >>",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>");

        PdfCatalogAction action = Assert.Single(PdfReadDocument.Open(pdf).CatalogActions);

        Assert.Equal("Open", action.Name);
        Assert.Equal("JavaScript", action.ActionType);
    }

    [Fact]
    public void PdfReadDocument_AppliesNodeBudgetSeparatelyToEachCatalogNameTree() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 5 0 R /JavaScript 6 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Target) [3 0 R /Fit]] >>",
            "<< /Names [(Open) 7 0 R] >>",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        Assert.Equal("Target", Assert.Single(document.NamedDestinations).Name);
        Assert.Equal("Open", Assert.Single(document.CatalogActions).Name);
    }

    [Fact]
    public void PdfReadDocument_IgnoresUnknownCatalogNameTreeAliasesBeforeTraversal() {
        string aliases = string.Join(
            " ",
            Enumerable.Range(0, 256).Select(static index => "/Unknown" + index + " 6 0 R"));
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 5 0 R " + aliases + " >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [(Target) [3 0 R /Fit]] >>",
            "<< /Kids [7 0 R] >>",
            "<< /Names [(Ignored) [3 0 R /Fit]] >>");
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfReadDocument document = PdfReadDocument.Open(pdf, options);

        Assert.Equal("Target", Assert.Single(document.NamedDestinations).Name);
    }

    [Fact]
    public void PdfReadDocument_BoundsEmbeddedFileNameTreeDepth() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles 5 0 R >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Kids [6 0 R] >>",
            "<< /Kids [7 0 R] >>",
            "<< /Names [] >>");
        (System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects, string trailer) = PdfSyntax.ParseObjects(pdf);
        var limits = new PdfReadLimits { MaxNameTreeDepth = 1 };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfAttachmentExtractor.ExtractAttachments(objects, trailer, limits));

        Assert.Equal(PdfReadLimitKind.NameTreeDepth, exception.Kind);
    }

    [Fact]
    public void PdfAttachmentExtractor_CountsOnlyIndirectNameTreeNodes() {
        byte[] pdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Kids [5 0 R] >> >> >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream(string.Empty),
            "<< /Names [] >>");
        (System.Collections.Generic.Dictionary<int, PdfIndirectObject> objects, string trailer) = PdfSyntax.ParseObjects(pdf);
        var limits = new PdfReadLimits { MaxNameTreeNodes = 1 };

        IReadOnlyList<PdfExtractedAttachment> attachments = PdfAttachmentExtractor.ExtractAttachments(objects, trailer, limits);

        Assert.Empty(attachments);
    }

    [Fact]
    public void PdfReadPage_BoundsBaseAndFallbackFontDecodingBeforeAllocation() {
        byte[] baseFontPdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            BuildStream("BT /F1 12 Tf (AB) Tj ET"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        byte[] fallbackFontPdf = BuildPdfObjects(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R >>",
            BuildStream("BT /Missing 12 Tf (AB) Tj ET"));
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 1 }
        };

        PdfReadLimitException baseException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(baseFontPdf, options).Pages[0].ExtractText());
        PdfReadLimitException fallbackException = Assert.Throws<PdfReadLimitException>(() =>
            PdfReadDocument.Open(fallbackFontPdf, options).Pages[0].ExtractText());

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, baseException.Kind);
        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, fallbackException.Kind);
    }

    [Fact]
    public void StandardEncoding_BoundsLigatureExpansionBeforeBuildingOutput() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfStandardEncoding.Decode(new byte[] { 174 }, maxOutputCharacters: 1));

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
    }

    [Fact]
    public void ContentStructureExtractor_RejectsHostileLeaderTextInLinearTime() {
        string hostile = "Label" + new string('-', 200_000) + "!";
        var spans = new[] { new PdfTextSpan(hostile, "F1", 12, 0, 100, hostile.Length) };
        var timer = Stopwatch.StartNew();

        StructuredPage page = ContentStructureExtractor.Extract(spans, new TextLayoutEngine.Options());

        timer.Stop();
        Assert.Empty(page.LeaderRows);
        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(5), "Leader parsing exceeded the linear-time test budget: " + timer.Elapsed + ".");
    }

    [Fact]
    public void ContentStructureExtractor_DoesNotTreatUnicodeDigitsAsAsciiPageNumbers() {
        const string text = "Contents.... ９";
        var spans = new[] { new PdfTextSpan(text, "F1", 12, 0, 100, text.Length) };

        StructuredPage page = ContentStructureExtractor.Extract(spans, new TextLayoutEngine.Options());

        Assert.Empty(page.Toc);
    }

    private static string BuildStream(string content, string dictionaryEntries = "") =>
        "<< " + dictionaryEntries + " /Length " + Encoding.ASCII.GetByteCount(content) + " >>\nstream\n" + content + "\nendstream";

    private static byte[] BuildPdfObjects(params string[] bodies) {
        var builder = new StringBuilder("%PDF-1.7\n");
        for (int index = 0; index < bodies.Length; index++) {
            builder.Append(index + 1).Append(" 0 obj\n").Append(bodies[index]).Append("\nendobj\n");
        }

        builder.Append("trailer\n<< /Root 1 0 R /Size ")
            .Append(bodies.Length + 1)
            .Append(" >>\nstartxref\n0\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
