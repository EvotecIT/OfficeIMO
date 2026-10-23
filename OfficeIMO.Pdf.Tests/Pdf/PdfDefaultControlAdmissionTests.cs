using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDefaultControlAdmissionTests {
    [Theory]
    [InlineData(false, 0x200E, false)]
    [InlineData(true, 0x200E, false)]
    [InlineData(false, 0x200B, false)]
    [InlineData(true, 0x200B, false)]
    [InlineData(false, 0x200E, true)]
    [InlineData(true, 0x200E, true)]
    public void DefaultControlsUseTheCompleteProviderResultAndTheSameMeasuredPens(bool cff, int control, bool named) {
        byte[] data = CreateControlFont(cff, blankProfile: 0, control);
        string text = "A" + char.ConvertFromUtf32(control) + "A";
        int successes = 0;
        var options = PdfTextShapingOptions.ForRendering("Control admission", PdfTextShapingMode.OpenTypeLigatures,
            providerShapedTextRecorder: (_, _, _, _) => successes++);
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data, "Control admission");
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Control admission") : null;

        PdfGlyphRun run = cff ? ot!.ShapeText(text, options) : tt!.ShapeText(text, options);

        AssertCompleteControlSource(text, run);
        Assert.Equal(1, successes);
        double measured = cff ? ot!.MeasureTextWidth(text, 10D, PdfTextShapingMode.OpenTypeLigatures)
            : tt!.MeasureTextWidth(text, 10D, PdfTextShapingMode.OpenTypeLigatures);
        Assert.Equal(run.TotalAdvanceWidth1000 * 10D / 1000D, measured, 8);
        PdfOptions pdfOptions = named
            ? new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Control admission", data))
            : new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Control admission");
        PdfTextShapingOptions diagnosticOptions = PdfTextDiagnostics.ForEncodingDiagnostics(pdfOptions, "Control admission");
        if (cff) ot!.ResetGlyphUsage(); else tt!.ResetGlyphUsage();
        var mappingsBefore = cff ? ot!.GetGlyphToUnicodeMappings().ToArray() : tt!.GetGlyphToUnicodeMappings().ToArray();
        Assert.Empty(cff ? PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, ot!, shapingOptions: diagnosticOptions)
            : PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, tt!, shapingOptions: diagnosticOptions));
        Assert.Empty(cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds());
        Assert.Equal(mappingsBefore, cff ? ot!.GetGlyphToUnicodeMappings().ToArray() : tt!.GetGlyphToUnicodeMappings().ToArray());
        PdfDocument document = PdfDocument.Create(pdfOptions).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Control admission");
            paragraph.Text(text);
        });
        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] pdf = document.ToBytes();
        // Paragraph layout intentionally treats ZWSP as an optional break in an
        // identifier. The complete provider run above still owns it exactly once.
        string publicText = control == 0x200B ? "AA" : text;
        Assert.Equal(publicText, PdfReadDocument.Open(pdf).ExtractText().Trim());
    }

    [Theory]
    [InlineData(false, 1, false)]
    [InlineData(true, 1, false)]
    [InlineData(false, 2, false)]
    [InlineData(true, 2, false)]
    [InlineData(false, 3, false)]
    [InlineData(true, 3, false)]
    [InlineData(false, 3, true)]
    [InlineData(true, 3, true)]
    public void DefaultControlsCannotFallThroughToMissingOrInkedCmapGlyphs(bool cff, int blankProfile, bool explicitManaged) {
        const int control = 0x200E;
        byte[] data = CreateControlFont(cff, blankProfile, control);
        const string text = "A\u200EA";
        IOfficeTextShapingProvider? provider = explicitManaged ? OfficeManagedTextShapingProvider.Instance : null;
        int successes = 0;
        var options = PdfTextShapingOptions.ForRendering("Unsafe control", PdfTextShapingMode.OpenTypeLigatures,
            provider, (_, _, _, _) => successes++);
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data, "Unsafe control");
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Unsafe control") : null;

        ArgumentException rendering = Assert.Throws<ArgumentException>(() => {
            if (cff) ot!.ShapeText(text, options); else tt!.ShapeText(text, options);
        });
        Assert.Equal("text", rendering.ParamName);
        Assert.Contains("verified blank glyphs", rendering.Message, StringComparison.Ordinal);
        ArgumentException measurement = Assert.Throws<ArgumentException>(() => {
            if (cff) ot!.MeasureTextWidth(text, 10D, PdfTextShapingMode.OpenTypeLigatures, provider);
            else tt!.MeasureTextWidth(text, 10D, PdfTextShapingMode.OpenTypeLigatures, provider);
        });
        Assert.Equal("text", measurement.ParamName);
        if (cff) ot!.ResetGlyphUsage(); else tt!.ResetGlyphUsage();
        PdfTextShapingOptions diagnosticOptions = PdfTextShapingOptions.ForDiagnostics("Unsafe control", "Unsafe control", PdfTextShapingMode.OpenTypeLigatures, provider);
        PdfGlyphRun diagnosticRun = cff ? ot!.ShapeText(text, diagnosticOptions) : tt!.ShapeText(text, diagnosticOptions);
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(diagnosticRun.Diagnostics);
        Assert.Equal(1, diagnostic.Index);
        Assert.Equal("U+200E", diagnostic.CodePoint);
        Assert.Equal("Unsafe control", diagnostic.Source);
        Assert.Equal(0, successes);
        Assert.Empty(diagnosticRun.Glyphs);
        Assert.Null(diagnosticRun.SourceShapingResult);
        Assert.False(diagnosticRun.IsAutomaticallyShaped);
        Assert.Empty(cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds());
    }

    [Fact]
    public void ImplicitDirectionalControlsCannotFallThroughToAnInkedCmapGlyph() {
        // A nominal Hebrew font with no GSUB/GPOS admits the visible letters.
        // Its actual space and LRM outlines have ink, so a control must decline.
        byte[] data = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, 0x05D0, 0x200E);
        var font = PdfTrueTypeFontProgram.Parse(data, "Directional control");
        Assert.Equal(2, font.ShapeText("\u05D0\u05D0",
            PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures)).Glyphs.Count);
        font.ResetGlyphUsage();
        const string text = "\u05D0\u200E\u05D0";
        int successes = 0;
        Assert.Throws<ArgumentException>(() => font.ShapeText(text,
            PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures,
                providerShapedTextRecorder: (_, _, _, _) => successes++)));
        PdfGlyphRun run = font.ShapeText(text,
            PdfTextShapingOptions.ForDiagnostics("Directional source", font.FontName, PdfTextShapingMode.OpenTypeLigatures));
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(run.Diagnostics);
        Assert.Equal("Directional source", diagnostic.Source);
        Assert.Equal(1, diagnostic.Index);
        Assert.Equal("U+200E", diagnostic.CodePoint);
        Assert.Empty(run.Glyphs);
        Assert.Null(run.SourceShapingResult);
        Assert.Equal(0, successes);
        Assert.Empty(font.GetUsedGlyphIds());
    }

    [Fact]
    public void UnqualifiedCarlitoFilteringStillDeclinesCompleteControls() {
        byte[] data = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
        const string text = "A\u200EA";
        var font = PdfTrueTypeFontProgram.Parse(data, "Unqualified filtering");
        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(
            text, font.FontName, data, false, font.UnitsPerEm, OfficeTextDirection.Auto, null, default,
            fontCollectionIndex: null, variationCoordinates: null, cloneFontData: true, applyDefaultLatinLigatures: true)));
        Assert.Throws<ArgumentException>(() => font.ShapeText(text,
            PdfTextShapingOptions.ForRendering(font.FontName, PdfTextShapingMode.OpenTypeLigatures)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LiteralControlDiagnosticsKeepTheirOriginalOwner(bool cff) {
        byte[] data = CreateControlFont(cff, blankProfile: 0, 0x200E);
        const string text = "A\0";
        PdfTextShapingOptions options = PdfTextShapingOptions.ForDiagnostics("Literal control", "Literal control", PdfTextShapingMode.OpenTypeLigatures);
        PdfGlyphRun run = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Literal control").ShapeText(text, options)
            : PdfTrueTypeFontProgram.Parse(data, "Literal control").ShapeText(text, options);
        // CFF's existing public raw diagnostic scan owns C0/DEL; unlike its
        // shape-only scalar fallback it materializes the expected diagnostics.
        PdfTextEncodingDiagnostic diagnostic = Assert.Single(cff
            ? PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, data, "Literal control")
            : run.Diagnostics);
        Assert.True(diagnostic.IsControlCharacter);
        Assert.Equal(1, diagnostic.Index);
        Assert.Equal("U+0000", diagnostic.CodePoint);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedControlPreflightUsesTheAuthoredFeatureDescriptor(bool named) {
        byte[] data = CreateControlFont(cff: true, blankProfile: 0, 0x200E);
        PdfOptions options = named
            ? new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Control features", data))
            : new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Control features");
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Control features");
            paragraph.FeatureSettings(OfficeTextFeatureSettings.Default.With("aalt", 1)).Text("A\u200EA");
        });
        Assert.NotEmpty(document.AnalyzeTextEncoding());
        Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GeneratedControlPreflightDoesNotBorrowManagedCoverageForCustomProviders(bool cff, bool named) {
        byte[] data = CreateControlFont(cff, blankProfile: 0, 0x200E);
        var provider = new UncalledControlProvider();
        PdfOptions options = (named
            ? new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Custom control", data))
            : new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Custom control"))
            .SetTextShapingProvider(provider);
        PdfTrueTypeFontProgram? tt = cff ? null : PdfTrueTypeFontProgram.Parse(data, "Custom control");
        PdfOpenTypeCffFontProgram? ot = cff ? PdfOpenTypeCffFontProgram.Parse(data, "Custom control") : null;
        PdfTextShapingOptions descriptor = PdfTextDiagnostics.ForEncodingDiagnostics(options, "Custom control");
        // These are the same internal raw coverage queries used by writer
        // measurement/encoding diagnostics. Ordinary covered text must not call
        // a provider either; generated LRM preflight alone would miss that path.
        Func<string, IReadOnlyList<PdfTextEncodingDiagnostic>> inspect = text => cff
            ? PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, ot!, shapingOptions: descriptor)
            : PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, tt!, shapingOptions: descriptor);
        Assert.Empty(inspect("AA"));
        Assert.NotEmpty(inspect("A\u200EA"));
        Assert.Equal(0, provider.Calls);
        Assert.Empty(cff ? ot!.GetUsedGlyphIds() : tt!.GetUsedGlyphIds());
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Custom control");
            paragraph.Text("A\u200EA");
        });
        Assert.NotEmpty(document.AnalyzeTextEncoding());
        Assert.Equal(0, provider.Calls);
        Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
        Assert.Equal(0, provider.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CffCompleteControlsRetainSourceAcrossLayoutAndNeededFallback(bool named) {
        const string family = "CFF layout controls";
        const string source = "A\u200EA\nA\u200EAB";
        byte[] primary = CreateControlFont(cff: true, blankProfile: 0, 0x200E);
        byte[] secondary = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'B');
        var report = new PdfConversionReport();
        PdfOptions options = (named
            ? new PdfOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, primary))
            : new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, primary, family))
            .SetTextShapingMode(PdfTextShapingMode.OpenTypeLigatures)
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("CFF needed B fallback", secondary)
            }))
            .ReportDiagnosticsTo(report);
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily(family);
            paragraph.Text(source);
        });

        Assert.Empty(document.AnalyzeTextEncoding());
        byte[] bytes = document.ToBytes();
        Assert.NotEmpty(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        // Only LF is a layout separator. Both original LRM controls and every
        // primary/fallback source character must survive exactly once.
        Assert.Equal("A\u200EAA\u200EAB", new string(extracted.Where(character => !char.IsWhiteSpace(character)).ToArray()));
        report.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LayoutFallbackCoverageCannotBorrowManagedNormalizationForScalarOrCustomProviders(bool customProvider) {
        const string source = "A\u0301\nB";
        byte[] primary = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'A', 0x00C1);
        byte[] secondary = OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(true, 'B');
        var provider = new UncalledControlProvider();
        PdfOptions options = new PdfOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, primary, "Unsupported normalization descriptor")
            .RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(new[] {
                new PdfEmbeddedFontFallbackCandidate("Descriptor B fallback", secondary)
            }))
            .SetTextShapingMode(customProvider ? PdfTextShapingMode.OpenTypeLigatures : PdfTextShapingMode.UnicodeScalar);
        if (customProvider) options.SetTextShapingProvider(provider);
        PdfDocument document = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text(source));

        var diagnostics = document.AnalyzeTextEncoding();
        Assert.NotEmpty(diagnostics);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Index < 2);
        Assert.Equal(0, provider.Calls);
        Assert.ThrowsAny<ArgumentException>(() => document.ToBytes());
        Assert.Equal(0, provider.Calls);
    }

    private sealed class UncalledControlProvider : IOfficeTextShapingProvider {
        internal int Calls { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Calls++;
            throw new InvalidOperationException("Encoding analysis must not execute arbitrary providers.");
        }
    }

    internal static void AssertCompleteControlSource(string text, PdfGlyphRun run) {
        Assert.NotNull(run.SourceShapingResult);
        Assert.Equal(text, run.ActualText);
        int consumed = 0;
        foreach (PdfGlyphInfo glyph in run.Glyphs.Where(glyph => glyph.UnicodeText.Length > 0).OrderBy(glyph => glyph.TextIndex)) {
            Assert.Equal(consumed, glyph.TextIndex);
            Assert.Equal(text.Substring(consumed, glyph.UnicodeText.Length), glyph.UnicodeText);
            consumed += glyph.UnicodeText.Length;
        }
        Assert.Equal(text.Length, consumed);
        Assert.Contains(run.Glyphs, glyph => glyph.AdvanceWidth1000 == 0);
        // An empty Unicode continuation can still paint a required substitution.
        // Only glyphs owned by an original control cluster must be blank.
        Assert.All(run.Glyphs.Where(glyph => glyph.LogicalClusterStart < text.Length
            && OfficeManagedTextShaper.IsNonPaintingControl(char.ConvertToUtf32(text, glyph.LogicalClusterStart))),
            glyph => {
                Assert.Equal(0, glyph.AdvanceWidth1000);
                Assert.Equal(0, glyph.OffsetX1000);
                Assert.Equal(0, glyph.OffsetY1000);
            });
    }

    // Existing real outlines and font tables remain unchanged. Only the cmap is
    // replaced: 0=real blank space; 1=no blank candidate; 2=inked space;
    // 3=inked control with no space. Profile3 used to paint the raw cmap glyph.
    private static byte[] CreateControlFont(bool cff, int blankProfile, int control) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
            : Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Drawing.Tests", "TestAssets", "OpenSans-Bold.ttf");
        Assert.NotNull(path);
        byte[] source = File.ReadAllBytes(path!);
        IOfficeFontProgram? program = cff ? OfficeOpenTypeCffFont.TryLoad(source, null, out _) : OfficeTrueTypeFont.TryLoad(source);
        Assert.NotNull(program);
        Assert.True(program!.TryGetGlyphMetrics('A', out int ink, out _));
        Assert.True(program.TryGetGlyphMetrics(' ', out int blank, out _));
        var mappings = new SortedDictionary<int, int> { ['A'] = ink };
        if (blankProfile == 0) mappings.Add(' ', blank);
        else if (blankProfile == 2) mappings.Add(' ', ink);
        else if (blankProfile == 3) mappings.Add(control, ink);
        var cmap = new byte[28 + mappings.Count * 12];
        Write16(cmap, 2, 1); Write16(cmap, 4, 3); Write16(cmap, 6, 10); Write32(cmap, 8, 12);
        Write16(cmap, 12, 12); Write32(cmap, 16, checked((uint)(cmap.Length - 12)));
        Write32(cmap, 24, checked((uint)mappings.Count));
        int group = 28;
        foreach (var mapping in mappings) {
            Write32(cmap, group, checked((uint)mapping.Key));
            Write32(cmap, group + 4, checked((uint)mapping.Key));
            Write32(cmap, group + 8, checked((uint)mapping.Value));
            group += 12;
        }
        int appended = (source.Length + 3) & ~3;
        var result = new byte[(appended + cmap.Length + 3) & ~3];
        Array.Copy(source, result, source.Length); Array.Copy(cmap, 0, result, appended, cmap.Length);
        int cmapRecord = TableRecord(result, "cmap");
        Write32(result, cmapRecord + 4, Checksum(cmap));
        Write32(result, cmapRecord + 8, checked((uint)appended));
        Write32(result, cmapRecord + 12, checked((uint)cmap.Length));
        int headRecord = TableRecord(result, "head"), head = checked((int)Read32(result, headRecord + 8));
        Write32(result, head + 8, 0);
        byte[] headBytes = result.Skip(head).Take(checked((int)Read32(result, headRecord + 12))).ToArray();
        Write32(result, headRecord + 4, Checksum(headBytes));
        Write32(result, head + 8, unchecked(0xB1B0AFBAU - Checksum(result)));
        return result;
    }

    private static int TableRecord(byte[] data, string tag) => Enumerable.Range(0, data[4] * 256 + data[5])
        .Select(index => 12 + index * 16).Single(record => Encoding.ASCII.GetString(data, record, 4) == tag);
    private static uint Checksum(byte[] bytes) {
        uint sum = 0;
        for (int index = 0; index < bytes.Length; index += 4) {
            uint value = 0;
            for (int offset = 0; offset < 4; offset++) value = (value << 8) | (index + offset < bytes.Length ? bytes[index + offset] : 0U);
            sum = unchecked(sum + value);
        }
        return sum;
    }
    private static uint Read32(byte[] bytes, int offset) => (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16
        | (uint)bytes[offset + 2] << 8 | bytes[offset + 3];
    private static void Write16(byte[] bytes, int offset, int value) { bytes[offset] = (byte)(value >> 8); bytes[offset + 1] = (byte)value; }
    private static void Write32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24); bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8); bytes[offset + 3] = (byte)value;
    }
}
