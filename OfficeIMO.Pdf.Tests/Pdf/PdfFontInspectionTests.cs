using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfFontInspectionTests {
    [Fact]
    public void FontInspectionDiagnosticCodes_PreserveStableLegacyValues() {
        Assert.Equal(0, (int)PdfFontInspectionDiagnosticCode.MissingBaseFont);
        Assert.Equal(1, (int)PdfFontInspectionDiagnosticCode.MissingToUnicode);
        Assert.Equal(2, (int)PdfFontInspectionDiagnosticCode.UnreadableToUnicode);
        Assert.Equal(3, (int)PdfFontInspectionDiagnosticCode.EmbeddedProgramUnavailable);
        Assert.Equal(4, (int)PdfFontInspectionDiagnosticCode.UnreadableEmbeddedOpenTypeProgram);
        Assert.Equal(5, (int)PdfFontInspectionDiagnosticCode.FontLimitExceeded);
        Assert.Equal(6, (int)PdfFontInspectionDiagnosticCode.ResourceReferenceLimitExceeded);
        Assert.Equal(7, (int)PdfFontInspectionDiagnosticCode.ResourceDepthExceeded);
        Assert.Equal(8, (int)PdfFontInspectionDiagnosticCode.CyclicResourceGraph);
    }

    [Fact]
    public void Fonts_ReportsSubsetEmbeddingToUnicodeAndNestedReferences() {
        byte[] pdf = BuildFontPdf();

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true
        });

        PdfFontInfo font = Assert.Single(inventory.Fonts);
        Assert.Equal(4, font.ObjectNumber);
        Assert.Equal(0, font.Generation);
        Assert.Equal("ABCDEF+DemoFont", font.BaseFontName);
        Assert.Equal("DemoFont", font.FamilyName);
        Assert.Equal("ABCDEF", font.SubsetTag);
        Assert.True(font.IsSubset);
        Assert.Equal("Type1", font.Subtype);
        Assert.Equal("WinAnsiEncoding", font.Encoding);
        Assert.True(font.HasToUnicode);
        Assert.True(font.HasReadableToUnicodeMap);
        Assert.Equal(1, font.ToUnicodeMappingCount);
        Assert.True(font.IsEmbedded);
        Assert.Equal("Type1", font.EmbeddedProgramSubtype);
        Assert.Equal(14, font.EmbeddedProgramEncodedLength);
        Assert.Equal(Encoding.ASCII.GetBytes("fake-font-data"), font.EmbeddedProgramBytes);
        Assert.Equal(2, font.References.Count);
        Assert.Contains(font.References, reference => reference.PageNumber == 1 && reference.ResourcePath == "Page 1/Font/F1");
        Assert.Contains(font.References, reference => reference.PageNumber == 1 && reference.ResourcePath == "Page 1/XObject/Fm1/Font/Nested");
        Assert.Empty(font.Diagnostics);
        Assert.Empty(inventory.Diagnostics);
        Assert.Equal(1, inventory.EmbeddedFontCount);
        Assert.Equal(1, inventory.SubsetFontCount);
        Assert.Equal(2, inventory.ResourceReferenceCount);
    }

    [Fact]
    public void Fonts_DoesNotRetainEmbeddedProgramBytesByDefault() {
        PdfFontInfo font = Assert.Single(PdfDocument.Load(BuildFontPdf()).Resources.Fonts().Fonts);

        Assert.True(font.IsEmbedded);
        Assert.Equal(14, font.EmbeddedProgramEncodedLength);
        Assert.Null(font.EmbeddedProgramBytes);
        Assert.Null(font.EmbeddedProgramDecodedLength);
    }

    [Fact]
    public void Fonts_InspectsEmbeddedTrueTypeProgramWithoutRetainingBytes() {
        byte[] fontBytes = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B');
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressEmbeddedFonts = true })
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Inspection")
            .Paragraph(paragraph => paragraph.Text("AB"))
            .ToBytes();

        PdfFontInfo font = Assert.Single(PdfDocument.Load(pdf).Resources.Fonts().Fonts, static candidate => candidate.IsEmbedded);
        PdfOpenTypeFontInfo program = Assert.IsType<PdfOpenTypeFontInfo>(font.EmbeddedOpenTypeInfo);

        Assert.True(program.IsTrueType);
        Assert.Equal(2, program.GlyphCount);
        Assert.Equal(3, program.UnicodeScalarCount);
        Assert.True(program.ContainsUnicodeScalar('A'));
        Assert.True(program.ContainsUnicodeScalar('B'));
        Assert.Null(font.EmbeddedProgramBytes);
        Assert.Null(font.EmbeddedProgramDecodedLength);
    }

    [Fact]
    public void Fonts_CanDisableEmbeddedProgramMetadataInspection() {
        byte[] fontBytes = ManagedTextShapingTestAssets.CreateFont(' ', 'A');
        byte[] pdf = PdfDocument.Create()
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "OfficeIMO Inspection")
            .Paragraph(paragraph => paragraph.Text("A"))
            .ToBytes();

        PdfFontInfo font = Assert.Single(PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            InspectEmbeddedProgramMetadata = false
        }).Fonts, static candidate => candidate.IsEmbedded);

        Assert.Null(font.EmbeddedOpenTypeInfo);
    }

    [Fact]
    public void Fonts_ReturnsDefensiveCopiesOfEmbeddedProgramBytes() {
        PdfFontInfo font = Assert.Single(PdfDocument.Load(BuildFontPdf()).Resources.Fonts(new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true
        }).Fonts);

        byte[] first = Assert.IsType<byte[]>(font.EmbeddedProgramBytes);
        first[0] = 0;

        Assert.Equal((byte)'f', Assert.IsType<byte[]>(font.EmbeddedProgramBytes)[0]);
    }

    [Fact]
    public void Fonts_ReportsMissingToUnicodeWithoutTreatingItAsUnreadable() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            StreamObject("BT /F1 12 Tf 10 200 Td (Hello) Tj ET"));

        PdfFontInfo font = Assert.Single(PdfDocument.Load(pdf).Resources.Fonts().Fonts);

        Assert.False(font.HasToUnicode);
        Assert.False(font.HasReadableToUnicodeMap);
        Assert.False(font.IsEmbedded);
        Assert.Contains(font.Diagnostics, diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.MissingToUnicode);
        Assert.Equal(1, PdfDocument.Load(pdf).Resources.Fonts().MissingToUnicodeFontCount);
    }

    [Fact]
    public void Fonts_StopsAtConfiguredReferenceLimitWithStructuredDiagnostic() {
        PdfFontInventory inventory = PdfDocument.Load(BuildFontPdf()).Resources.Fonts(new PdfFontInspectionOptions {
            MaxResourceReferences = 1
        });

        Assert.Equal(1, inventory.ResourceReferenceCount);
        PdfFontInspectionDiagnostic diagnostic = Assert.Single(inventory.Diagnostics);
        Assert.Equal(PdfFontInspectionDiagnosticCode.ResourceReferenceLimitExceeded, diagnostic.Code);
        Assert.Equal(1, diagnostic.PageNumber);
        Assert.Equal("Page 1/XObject/Fm1/Font/Nested", diagnostic.ResourcePath);
    }

    [Fact]
    public void Fonts_StopsDecodingWhenAggregateEmbeddedProgramLimitIsReached() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /FontOne /FontDescriptor 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /FontTwo /FontDescriptor 8 0 R >>",
            StreamObject(string.Empty),
            "<< /Type /FontDescriptor /FontName /FontOne /FontFile 9 0 R >>",
            "<< /Type /FontDescriptor /FontName /FontTwo /FontFile 10 0 R >>",
            StreamObject("font-one", "/Length1 8"),
            StreamObject("font-two", "/Length1 8"));

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            IncludeEmbeddedProgramBytes = true,
            MaxEmbeddedProgramBytes = 16,
            MaxTotalDecodedFontBytes = 8
        });

        Assert.Equal(2, inventory.FontCount);
        Assert.Equal(Encoding.ASCII.GetBytes("font-one"), inventory.Fonts[0].EmbeddedProgramBytes);
        Assert.Null(inventory.Fonts[1].EmbeddedProgramBytes);
        Assert.Contains(
            inventory.Fonts[1].Diagnostics,
            diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.EmbeddedProgramTotalLimitExceeded);
    }

    [Fact]
    public void Fonts_DistinguishesToUnicodePerMapLimitFromUnreadableContent() {
        const string toUnicode = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 beginbfchar\n<41> <0041>\nendbfchar\nendcmap\nend\nend";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>",
            StreamObject(string.Empty),
            StreamObject(toUnicode));

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxToUnicodeBytes = 16,
            MaxTotalDecodedFontBytes = 1_024
        });

        PdfFontInfo font = Assert.Single(inventory.Fonts);
        Assert.False(font.HasReadableToUnicodeMap);
        Assert.Contains(font.Diagnostics, diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ToUnicodeLimitExceeded);
        Assert.Equal(1, inventory.ToUnicodeLimitExceededFontCount);
        Assert.Equal(0, inventory.ToUnicodeTotalLimitExceededFontCount);
        Assert.Equal(0, inventory.UnreadableToUnicodeFontCount);
    }

    [Fact]
    public void Fonts_DistinguishesAggregateToUnicodeLimitFromUnreadableContent() {
        const string toUnicode = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 beginbfchar\n<41> <0041>\nendbfchar\nendcmap\nend\nend";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Courier /ToUnicode 8 0 R >>",
            StreamObject(string.Empty),
            StreamObject(toUnicode),
            StreamObject(toUnicode));

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxToUnicodeBytes = 1_024,
            MaxTotalDecodedFontBytes = Encoding.ASCII.GetByteCount(toUnicode)
        });

        Assert.True(inventory.Fonts[0].HasReadableToUnicodeMap);
        Assert.False(inventory.Fonts[1].HasReadableToUnicodeMap);
        Assert.Contains(
            inventory.Fonts[1].Diagnostics,
            diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ToUnicodeTotalLimitExceeded);
        Assert.Equal(0, inventory.ToUnicodeLimitExceededFontCount);
        Assert.Equal(1, inventory.ToUnicodeTotalLimitExceededFontCount);
        Assert.Equal(0, inventory.UnreadableToUnicodeFontCount);
    }

    [Fact]
    public void Fonts_BoundsRepeatedMalformedToUnicodeDecodeAttempts() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R /F2 5 0 R /F3 6 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 8 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Courier /ToUnicode 9 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /TimesRoman /ToUnicode 10 0 R >>",
            StreamObject(string.Empty),
            StreamObject("bad", "/Filter /DCTDecode"),
            StreamObject("bad", "/Filter /DCTDecode"),
            StreamObject("bad", "/Filter /DCTDecode"));

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxToUnicodeBytes = 16,
            MaxTotalDecodedFontBytes = 32
        });

        Assert.Equal(2, inventory.UnreadableToUnicodeFontCount);
        Assert.Equal(1, inventory.ToUnicodeTotalLimitExceededFontCount);
        Assert.Contains(
            inventory.Fonts[2].Diagnostics,
            diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ToUnicodeTotalLimitExceeded);
    }

    [Fact]
    public void Fonts_VisitsSharedFormResourceContextOnlyOnce() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /A 4 0 R /B 5 0 R >> >> /Contents 6 0 R >>",
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Shared 7 0 R >> >>"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Shared 7 0 R >> >>"),
            StreamObject(string.Empty),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /Font << /F1 8 0 R >> >>"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxFormResourceTraversals = 4
        });

        Assert.Single(inventory.Fonts);
        Assert.Equal(2, inventory.Fonts[0].References.Count);
        Assert.Contains(inventory.Fonts[0].References, reference => reference.ResourcePath == "Page 1/XObject/A/XObject/Shared/Font/F1");
        Assert.Contains(inventory.Fonts[0].References, reference => reference.ResourcePath == "Page 1/XObject/B/XObject/Shared/Font/F1");
        Assert.DoesNotContain(
            inventory.Diagnostics,
            diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.FormResourceTraversalLimitExceeded);
    }

    [Fact]
    public void Fonts_RevisitsSharedFormWhenLaterPathHasMoreDepthAvailable() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /Deep 4 0 R /Shallow 5 0 R >> >> /Contents 6 0 R >>",
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Middle 7 0 R >> >>"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Shared 8 0 R >> >>"),
            StreamObject(string.Empty),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Shared 8 0 R >> >>"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /Leaf 9 0 R >> >>"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /Font << /F1 10 0 R >> >>"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxResourceDepth = 3,
            MaxFormResourceTraversals = 10
        });

        PdfFontInfo font = Assert.Single(inventory.Fonts);
        PdfFontResourceReference reference = Assert.Single(font.References);
        Assert.Equal("Page 1/XObject/Shallow/XObject/Shared/XObject/Leaf/Font/F1", reference.ResourcePath);
        Assert.Contains(inventory.Diagnostics, diagnostic => diagnostic.Code == PdfFontInspectionDiagnosticCode.ResourceDepthExceeded);
    }

    [Fact]
    public void Fonts_StopsAtConfiguredFormResourceTraversalLimit() {
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /XObject << /A 4 0 R /B 5 0 R >> >> /Contents 6 0 R >>",
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>"),
            StreamObject(string.Empty));

        PdfFontInventory inventory = PdfDocument.Load(pdf).Resources.Fonts(new PdfFontInspectionOptions {
            MaxFormResourceTraversals = 1
        });

        PdfFontInspectionDiagnostic diagnostic = Assert.Single(inventory.Diagnostics);
        Assert.Equal(PdfFontInspectionDiagnosticCode.FormResourceTraversalLimitExceeded, diagnostic.Code);
        Assert.Equal("Page 1/XObject/B", diagnostic.ResourcePath);
    }

    [Fact]
    public void TryFonts_UsesLogicalContentPermissionGate() {
        PdfOperationResult<PdfFontInventory> result = PdfDocument.Load(BuildFontPdf()).Reader.TryFonts();

        Assert.True(result.Succeeded);
        Assert.Equal(PdfPreflightCapability.ReadLogicalObjects, result.Capability);
        Assert.Equal(1, result.RequireValue().FontCount);
    }

    [Fact]
    public void Fonts_PageSelector_UsesDocumentRelativeSelectionContract() {
        PdfDocument source = PdfDocument.Load(BuildFontPdf());

        PdfFontInventory inventory = source.Resources.Fonts(PdfPageSelector.Parse("last"));
        PdfOperationResult<PdfFontInventory> attempt = source.Reader.TryFonts(PdfPageSelector.Parse("1"));

        Assert.Equal(1, inventory.FontCount);
        Assert.True(attempt.Succeeded);
        Assert.Equal(1, attempt.RequireValue().FontCount);
    }

    private static byte[] BuildFontPdf() {
        const string toUnicode = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n1 beginbfchar\n<41> <0041>\nendbfchar\nendcmap\nend\nend";
        return BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> /XObject << /Fm1 7 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+DemoFont /Encoding /WinAnsiEncoding /FirstChar 65 /LastChar 65 /Widths [600] /FontDescriptor 5 0 R /ToUnicode 8 0 R >>",
            "<< /Type /FontDescriptor /FontName /ABCDEF+DemoFont /Flags 32 /FontBBox [0 -200 1000 900] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile 9 0 R >>",
            StreamObject("BT /F1 12 Tf 10 200 Td (A) Tj ET /Fm1 Do"),
            StreamObject(string.Empty, "/Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /Nested 4 0 R >> >>"),
            StreamObject(toUnicode),
            StreamObject("fake-font-data", "/Length1 14"));
    }

    private static string StreamObject(string content, string additionalDictionary = "") {
        int length = Encoding.ASCII.GetByteCount(content);
        string suffix = string.IsNullOrWhiteSpace(additionalDictionary) ? string.Empty : " " + additionalDictionary;
        return "<< /Length " + length.ToString(CultureInfo.InvariantCulture) + suffix + " >>\nstream\n" + content + "\nendstream";
    }

    private static byte[] BuildPdf(params string[] objects) {
        var builder = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>(objects.Length);
        for (int i = 0; i < objects.Length; i++) {
            offsets.Add(Encoding.ASCII.GetByteCount(builder.ToString()));
            builder.Append(i + 1).Append(" 0 obj\n").Append(objects[i]).Append("\nendobj\n");
        }

        int xrefOffset = Encoding.ASCII.GetByteCount(builder.ToString());
        builder.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        for (int i = 0; i < offsets.Count; i++) {
            builder.Append(offsets[i].ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        }
        builder.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1).Append(" >>\nstartxref\n")
            .Append(xrefOffset.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
