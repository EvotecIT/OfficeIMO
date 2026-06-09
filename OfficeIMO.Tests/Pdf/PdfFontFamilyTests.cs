using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFontFamilyTests {
    [Fact]
    public void PdfEmbeddedFontFamily_SnapshotsFaceBytesForReusableRegistration() {
        byte[] regular = { 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        byte[] bold = { 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0 };
        var family = new PdfEmbeddedFontFamily("Snapshot Family", regular, bold);
        regular[0] = 255;
        bold[5] = 255;

        PdfOptions options = new PdfOptions().UseFontFamily(family);
        byte[] readback = family.Regular;
        readback[0] = 127;

        Assert.Equal("Snapshot Family", family.FamilyName);
        Assert.Equal(0, family.Regular[0]);
        Assert.Equal(1, family.Bold![5]);
        Assert.Equal("Snapshot Family-Regular", options.EmbeddedFonts[PdfStandardFont.Helvetica].FontName);
        Assert.Equal("Snapshot Family-Bold", options.EmbeddedFonts[PdfStandardFont.HelveticaBold].FontName);
        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.FooterFont);
        Assert.Throws<ArgumentException>(() => new PdfEmbeddedFontFamily("Broken", Array.Empty<byte>()));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemLoadsInstalledTrueTypeFamily() {
        if (!TryFindInstalledSystemFontFamily(out PdfEmbeddedFontFamily? family)) {
            return;
        }

        PdfOptions options = new PdfOptions().UseFontFamily(family);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily(family)
            .Paragraph(paragraph => paragraph
                .Text("System regular ")
                .Bold("system bold ")
                .Italic("system italic ")
                .Text("system done"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.NotEmpty(family.Regular);
        Assert.True(family.Bold != null || family.Italic != null || family.BoldItalic != null);
        Assert.Contains(PdfStandardFont.Helvetica, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaBold, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaOblique, options.EmbeddedFonts.Keys);
        Assert.Contains(PdfStandardFont.HelveticaBoldOblique, options.EmbeddedFonts.Keys);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("System regular system bold system italic system done", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_FromSystemThrowsWhenFamilyIsMissing() {
        Assert.Throws<FileNotFoundException>(() =>
            PdfEmbeddedFontFamily.FromSystem("OfficeIMO Missing Font Family 404"));
    }

    [Fact]
    public void PdfOptions_UseOfficeFontFamilyFallsBackThroughFamilyListToStandardFont() {
        PdfOptions options = new PdfOptions()
            .UseOfficeFontFamily("OfficeIMO Missing Display, Consolas, monospace", embedSystemFont: false);

        Assert.Equal(PdfStandardFont.Courier, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Courier, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Courier, options.FooterFont);
    }

    [Fact]
    public void PdfOptions_TryUseDefaultDocumentFontFallbackEmbedsUnicodeCapableGeneratedText() {
        var options = new PdfOptions {
            CompressContentStreams = false
        };
        if (!options.TryUseDefaultDocumentFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text(polish))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.True(options.HasEmbeddedStandardFontFamily(options.DefaultFont));
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfOptions_TryRegisterDefaultDocumentMonospaceFontFallbackEmbedsCourierSlotWhenAvailable() {
        var options = new PdfOptions();
        if (!options.TryRegisterDefaultDocumentMonospaceFontFallback(requireEmbeddedFont: true)) {
            return;
        }

        Assert.True(options.HasEmbeddedStandardFontFamily(PdfStandardFont.Courier));
    }

    [Fact]
    public void PdfOptions_RegisterFontFamilyEmbedsSemanticSlotWithoutChangingDefaults() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions {
            CompressContentStreams = false
        }.RegisterFontFamily(
            PdfStandardFont.TimesRoman,
            PdfEmbeddedFontFamily.FromFiles("OfficeIMO Semantic Serif", fontPath));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph
                .Font(PdfStandardFont.TimesRoman)
                .Text("Semantic serif Łódź"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.HeaderFont);
        Assert.Equal(PdfStandardFont.Helvetica, options.FooterFont);
        Assert.Contains("/BaseFont /OfficeIMOSemanticSerif-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("Semantic serif Łódź", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesMatchesRenamedTrueTypeMetadata() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string renamedPath = Path.Combine(tempDir, "officeimo-renamed-system-face.ttf");
            File.Copy(fontPath, renamedPath);

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { renamedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.True(found);
            Assert.NotNull(family);
            Assert.NotEmpty(family!.Regular);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesSkipsReadableMetadataMismatchBeforeFilenameFallback() {
        if (!TryFindSingleInstalledRegularFontFace(out _, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string renamedPath = Path.Combine(tempDir, "OfficeIMOFilenameTrap-Regular.ttf");
            File.Copy(fontPath, renamedPath);

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                "OfficeIMO Filename Trap",
                new[] { renamedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.False(found);
            Assert.Null(family);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_MetadataMatchingDoesNotUseFileNameAliases() {
        Assert.False(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Times", "Times New Roman"));
        Assert.False(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Courier", "Courier New"));
        Assert.True(PdfEmbeddedFontFamily.IsMetadataFamilyNameMatch("Times New Roman", "Times New Roman"));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_MetadataStyleScoringPrefersExactBoldItalic() {
        Assert.True(
            PdfEmbeddedFontFamily.GetMetadataStyleScore("Bold Italic") >
            PdfEmbeddedFontFamily.GetMetadataStyleScore("SemiBold Italic"));
        Assert.True(
            PdfEmbeddedFontFamily.GetMetadataStyleScore("Bold") >
            PdfEmbeddedFontFamily.GetMetadataStyleScore("SemiBold"));
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesMatchesTrueTypeCollectionFace() {
        if (!TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath)) {
            return;
        }

        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string collectionPath = Path.Combine(tempDir, "officeimo-system-face.ttc");
            File.WriteAllBytes(collectionPath, CreateSingleFaceTrueTypeCollection(File.ReadAllBytes(fontPath)));

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                familyName,
                new[] { collectionPath },
                out PdfEmbeddedFontFamily? family);

            Assert.True(found);
            Assert.NotNull(family);
            Assert.NotEmpty(family!.Regular);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_TryFromSystemFontFilesSkipsMalformedTrueTypeFiles() {
        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.Pdf.Fonts." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try {
            string malformedPath = Path.Combine(tempDir, "OfficeIMOMalformed-Regular.ttf");
            File.WriteAllBytes(malformedPath, new byte[] { 0, 1, 0, 0, 0, 255, 255, 255 });

            bool found = PdfEmbeddedFontFamily.TryFromSystemFontFiles(
                "OfficeIMO Malformed",
                new[] { malformedPath },
                out PdfEmbeddedFontFamily? family);

            Assert.False(found);
            Assert.Null(family);
        } finally {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    [Fact]
    public void PdfEmbeddedFontFamily_GetSystemFontRootsIncludesWindowsPerUserFonts() {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData)) {
            return;
        }

        string expected = Path.Combine(localAppData, "Microsoft", "Windows", "Fonts");
        Assert.Contains(
            PdfEmbeddedFontFamily.GetSystemFontRoots(),
            root => string.Equals(root, expected, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PdfDocument_UseFontFamilyObjectReusesTrueTypeFamilyForGeneratedText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Object Font", fontData);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily(family)
            .Header(header => header.Text("Object font header"))
            .Paragraph(paragraph => paragraph
                .Text("Object regular ")
                .Bold("object bold ")
                .Italic("object italic"))
            .Footer(footer => footer.Text("Object font footer {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CIDToGIDMap /Identity", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Bold", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /OfficeIMOObjectFont-Italic", raw, StringComparison.Ordinal);
        Assert.Contains("/Length1 " + fontData.Length.ToString(CultureInfo.InvariantCulture), raw, StringComparison.Ordinal);
        Assert.Contains("Object regular object bold object italic", text, StringComparison.Ordinal);
        Assert.Contains("Object font header", text, StringComparison.Ordinal);
        Assert.Contains("Object font footer 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_UseFontFamilyScopesFamilyToComposedPageHeaderFooterAndBody() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var family = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Page Font", fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.Text("Before page font"))
            .Page(page => page
                .UseFontFamily(family)
                .Header(header => header.Text("Page font header"))
                .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page font body"))))
                .Footer(footer => footer.Text("Page font footer {page}/{pages}")))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOPageFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("Before page font", text, StringComparison.Ordinal);
        Assert.Contains("Page font header", text, StringComparison.Ordinal);
        Assert.Contains("Page font body", text, StringComparison.Ordinal);
        Assert.Contains("Page font footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_UseFontFamilyScopesFamilyToTextFieldAppearance() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var family = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Page Form Font", fontPath);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Page(page => page
                .UseFontFamily(family)
                .Content(content => content.Item(item => item.TextField("Scoped.Name", value: "Lodz"))))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        int embeddedFontObjectId = FindObjectIdBefore(raw, "/Subtype /Type0");

        Assert.Contains("/BaseFont /OfficeIMOPageFormFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/Resources << /Font << /Helv " + embeddedFontObjectId.ToString(CultureInfo.InvariantCulture) + " 0 R >>", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultTextStyle_FontFamilyDoesNotRewriteExistingHeaderOrFooterFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("OfficeIMO Style Font", fontData);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Header(header => header.Font(PdfStandardFont.TimesRoman).Text("Times header"))
            .Footer(footer => footer.Font(PdfStandardFont.Courier).Text("Courier footer"))
            .DefaultTextStyle(style => style.FontFamily(family).FontSize(12))
            .Paragraph(paragraph => paragraph.Text("Styled body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/BaseFont /OfficeIMOStyleFont-Regular", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Times-Roman", raw, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Courier", raw, StringComparison.Ordinal);
        Assert.Contains("Times header", text, StringComparison.Ordinal);
        Assert.Contains("Styled body", text, StringComparison.Ordinal);
        Assert.Contains("Courier footer", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_UseFontFamilyWritesUnicodeGlyphsAndToUnicodeExtraction() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        const string polish = "Zażółć gęślą jaźń Łódź";
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Unicode Font", fontPath)
            .Header(header => header.Text("Nagłówek Łódź"))
            .Paragraph(paragraph => paragraph
                .Text(polish + " ")
                .Bold("Śląsk ")
                .Italic("źrebak"))
            .Footer(footer => footer.Text("Stopka Łódź {page}/{pages}"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string text = PdfReadDocument.Load(bytes).ExtractText();

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
        Assert.Contains(polish, text, StringComparison.Ordinal);
        Assert.Contains("Nagłówek Łódź", text, StringComparison.Ordinal);
        Assert.Contains("Śląsk źrebak", text, StringComparison.Ordinal);
        Assert.Contains("Stopka Łódź 1/1", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_UsesEmbeddedFontCoverageForGeneratedText() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Font", fontPath);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(
            "Zażółć gęślą jaźń",
            options,
            PdfStandardFont.Helvetica,
            "body");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfTextDiagnostics_ReportsEmbeddedFontMissingGlyphAsConversionWarning() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Font", fontPath);
        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedText(
            "Snowman \u2603",
            options,
            PdfStandardFont.Helvetica,
            "body",
            "PdfParagraph[0]");

        if (diagnostics.Count == 0) {
            return;
        }

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnostics);
        PdfConversionWarning warning = diagnostic.ToConversionWarning("OfficeIMO.Tests");

        Assert.Equal("unsupported-text-glyph", diagnostic.Code);
        Assert.Equal("U+2603", diagnostic.CodePoint);
        Assert.Contains("embedded TrueType font", diagnostic.Encoding, StringComparison.Ordinal);
        Assert.Contains("fallback", diagnostic.Remediation, StringComparison.Ordinal);
        Assert.Equal("PdfParagraph[0]", diagnostic.Location);
        Assert.Equal(diagnostic.Encoding, warning.Details["encoding"]);
        Assert.Equal(diagnostic.Remediation, warning.Details["remediation"]);
        Assert.Equal(diagnostic.Location, warning.Details["location"]);
    }

    [Fact]
    public void PdfDocument_EmbeddedFontMissingGlyphThrowsDiagnosticBackedException() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        string missingGlyph = char.ConvertFromUtf32(0x10FFFF);
        ArgumentException exception = Assert.ThrowsAny<ArgumentException>(() =>
            PdfDocument.Create()
                .UseFontFamily("OfficeIMO Diagnostic Font", fontPath)
                .Paragraph(paragraph => paragraph.Text("Missing " + missingGlyph))
                .ToBytes());

        Assert.Contains("U+10FFFF", exception.Message, StringComparison.Ordinal);
        Assert.Contains("embedded TrueType font", exception.Message, StringComparison.Ordinal);
        Assert.Contains("fallback", exception.Message, StringComparison.Ordinal);
        Assert.Equal("unsupported-text-glyph", exception.Data["code"]);
        Assert.Equal("PdfParagraph", exception.Data["source"]);
        Assert.Equal("PdfParagraph[0].Run[0]", exception.Data["location"]);
        Assert.Equal(0, exception.Data["runIndex"]);
        Assert.Equal("U+10FFFF", exception.Data["codePoint"]);
        Assert.Equal(1, exception.Data["diagnosticsCount"]);
        Assert.Contains("embedded TrueType font", (string)exception.Data["encoding"]!, StringComparison.Ordinal);
        Assert.Contains("fallback", (string)exception.Data["remediation"]!, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTextDiagnostics_UsesStyledEmbeddedFontSlotsForRuns() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .UseFontFamily("OfficeIMO Diagnostic Family", fontPath);

        IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics = PdfTextDiagnostics.AnalyzeGeneratedTextRuns(
            new[] {
                TextRun.Bolded("Bold Łódź"),
                TextRun.Italicized("Italic Łódź"),
                TextRun.BoldItalic("Bold italic Łódź")
            },
            options,
            PdfStandardFont.Helvetica,
            "body");

        Assert.Empty(diagnostics);
    }

    [Fact]
    public void PdfDocument_UseFontFamilyEncodesTextWatermarkWithEmbeddedGlyphs() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .UseFontFamily("OfficeIMO Watermark Font", fontPath)
            .Watermark("DRAFT", fontSize: 32, opacity: 0.25)
            .Paragraph(paragraph => paragraph.Text("Watermark font proof"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("<4452414654> Tj", raw, StringComparison.Ordinal);
        Assert.Contains("/CMapName /OfficeIMO-Identity-Glyph-UCS", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_UseFontFamilyWrapsLongNonBmpTokensWithoutSplittingSurrogates() {
        foreach (string fontPath in EnumerateLocalNonBmpTrueTypeFonts()) {
            byte[] bytes;
            try {
                bytes = PdfDocument.Create(new PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 120,
                    MarginLeft = 24,
                    MarginRight = 24
                })
                .UseFontFamily("OfficeIMO NonBmp Font", fontPath)
                .Paragraph(paragraph => paragraph.Text("AAAAAAAAAAAA😀BBBBBBBBBBBB"))
                .ToBytes();
            } catch (ArgumentException exception) when (exception.Message.Contains("embedded TrueType font", StringComparison.Ordinal)) {
                continue;
            }

            string raw = Encoding.ASCII.GetString(bytes);

            Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
            Assert.Contains("/Encoding /Identity-H", raw, StringComparison.Ordinal);
            return;
        }
    }

    private static IEnumerable<string> EnumerateLocalNonBmpTrueTypeFonts() {
        string windowsFonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        string[] candidates = {
            Path.Combine(windowsFonts, "seguiemj.ttf"),
            Path.Combine(windowsFonts, "seguisym.ttf"),
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansSymbols2-Regular.ttf"
        };

        foreach (string candidate in candidates) {
            if (File.Exists(candidate)) {
                yield return candidate;
            }
        }
    }

    private static int FindObjectIdBefore(string raw, string marker) {
        int markerIndex = raw.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(markerIndex >= 0, "Could not find marker '" + marker + "'.");

        int objectSuffixIndex = raw.LastIndexOf(" obj", markerIndex, StringComparison.Ordinal);
        Assert.True(objectSuffixIndex >= 0, "Could not find object header before marker '" + marker + "'.");

        int lineStart = raw.LastIndexOf('\n', objectSuffixIndex);
        string objectHeader = raw.Substring(lineStart + 1, objectSuffixIndex - lineStart - 1);
        string[] parts = objectHeader.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        Assert.True(parts.Length >= 2, "Could not parse object header '" + objectHeader + "'.");
        return int.Parse(parts[0], CultureInfo.InvariantCulture);
    }

    private static bool TryFindInstalledSystemFontFamily(out PdfEmbeddedFontFamily? family) {
        string[] candidates = {
            "Arial",
            "DejaVu Sans",
            "Liberation Sans",
            "Segoe UI"
        };

        foreach (string candidate in candidates) {
            if (PdfEmbeddedFontFamily.TryFromSystem(candidate, out family) &&
                family != null &&
                (family.Bold != null || family.Italic != null || family.BoldItalic != null)) {
                return true;
            }
        }

        family = null;
        return false;
    }

    private static bool TryFindSingleInstalledRegularFontFace(out string familyName, out string fontPath) {
        string[] candidates = {
            "Arial",
            "DejaVu Sans",
            "Liberation Sans",
            "Segoe UI",
            "Microsoft Sans Serif"
        };

        foreach (string path in EnumerateInstalledTrueTypeFonts()) {
            foreach (string candidate in candidates) {
                if (PdfEmbeddedFontFamily.TryFromSystemFontFiles(candidate, new[] { path }, out PdfEmbeddedFontFamily? family) &&
                    family != null &&
                    family.Bold == null &&
                    family.Italic == null &&
                    family.BoldItalic == null) {
                    familyName = candidate;
                    fontPath = path;
                    return true;
                }
            }
        }

        familyName = string.Empty;
        fontPath = string.Empty;
        return false;
    }

    private static IEnumerable<string> EnumerateInstalledTrueTypeFonts() {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrWhiteSpace(windows)) {
            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(windows, "Fonts"))) {
                yield return font;
            }
        }

        foreach (string font in EnumerateTrueTypeFonts("/usr/share/fonts")) {
            yield return font;
        }

        foreach (string font in EnumerateTrueTypeFonts("/usr/local/share/fonts")) {
            yield return font;
        }

        string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrWhiteSpace(userProfile)) {
            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, ".local", "share", "fonts"))) {
                yield return font;
            }

            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, ".fonts"))) {
                yield return font;
            }

            foreach (string font in EnumerateTrueTypeFonts(Path.Combine(userProfile, "Library", "Fonts"))) {
                yield return font;
            }
        }

        foreach (string font in EnumerateTrueTypeFonts("/Library/Fonts")) {
            yield return font;
        }

        foreach (string font in EnumerateTrueTypeFonts("/System/Library/Fonts")) {
            yield return font;
        }
    }

    private static IEnumerable<string> EnumerateTrueTypeFonts(string root) {
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)) {
            yield break;
        }

        string[] files;
        try {
            files = Directory.GetFiles(root, "*.ttf", SearchOption.AllDirectories);
        } catch (IOException) {
            yield break;
        } catch (UnauthorizedAccessException) {
            yield break;
        }

        foreach (string file in files) {
            yield return file;
        }
    }

    private static byte[] CreateSingleFaceTrueTypeCollection(byte[] fontData) {
        int fontOffset = 16;
        int tableCount = ReadUInt16(fontData, 4);
        int sourceDirectoryLength = 12 + tableCount * 16;
        int collectionLength = fontOffset + fontData.Length;
        byte[] collection = new byte[collectionLength];

        collection[0] = (byte)'t';
        collection[1] = (byte)'t';
        collection[2] = (byte)'c';
        collection[3] = (byte)'f';
        WriteUInt32(collection, 4, 0x00010000);
        WriteUInt32(collection, 8, 1);
        WriteUInt32(collection, 12, (uint)fontOffset);
        Array.Copy(fontData, 0, collection, fontOffset, fontData.Length);

        for (int i = 0; i < tableCount; i++) {
            int recordOffset = fontOffset + 12 + i * 16;
            uint sourceOffset = ReadUInt32(fontData, 12 + i * 16 + 8);
            WriteUInt32(collection, recordOffset + 8, (uint)(fontOffset + sourceOffset));
        }

        Assert.True(sourceDirectoryLength <= fontData.Length);
        return collection;
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
