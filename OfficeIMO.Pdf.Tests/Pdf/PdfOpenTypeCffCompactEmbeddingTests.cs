using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOpenTypeCffCompactEmbeddingTests {
    [Fact]
    public void PdfOpenTypeCffFontProgram_BuildsCompactOpenTypeFontFileWithoutLayoutTables() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] full = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram program = PdfOpenTypeCffFontProgram.Parse(full, "OfficeIMO Source Serif CFF");
        byte[] compact = program.BuildCompactOpenTypeFontFile();
        PdfOpenTypeFontInfo compactInfo = PdfOpenTypeFontInspector.Inspect(compact, "OfficeIMO Source Serif CFF");
        IReadOnlyList<string> compactTables = ReadTableTags(compact);

        Assert.True(compact.Length < full.Length);
        Assert.True(compactInfo.IsOpenTypeCff);
        Assert.True(compactInfo.GlyphCount > 1000);
        Assert.True(compactInfo.CffTableLength > 1000);
        Assert.False(compactInfo.HasGlyphSubstitutionTable);
        Assert.False(compactInfo.HasGlyphPositioningTable);
        Assert.Contains("CFF ", compactTables);
        Assert.Contains("cmap", compactTables);
        Assert.Contains("hmtx", compactTables);
        Assert.DoesNotContain("GSUB", compactTables);
        Assert.DoesNotContain("GPOS", compactTables);
        Assert.DoesNotContain("DSIG", compactTables);
    }

    [Fact]
    public void PdfDocument_EmbedStandardFontWritesCompactOpenTypeCffFontFile3Output() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] fontData = File.ReadAllBytes(fontPath!);
        var report = new PdfConversionReport();
        var options = new PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Compact CFF Łódź"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Load(bytes).ExtractText();
        int embeddedLength = Assert.Single(ExtractLength1Values(raw));
        PdfConversionWarning warning = Assert.Single(report.Warnings, item => item.Code == "opentype-cff-charstrings-not-subset");

        Assert.Contains("Compact CFF Łódź", extracted, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /OpenType", raw, StringComparison.Ordinal);
        Assert.InRange(embeddedLength, 1, fontData.Length - 1);
        Assert.Contains("charstrings are still kept intact", warning.Message, StringComparison.Ordinal);
        Assert.Equal(fontData.Length.ToString(CultureInfo.InvariantCulture), warning.Details["fontFileLength"]);
        Assert.Equal(embeddedLength.ToString(CultureInfo.InvariantCulture), warning.Details["embeddedFontFileLength"]);
        Assert.Equal("compact-opentype-cff", warning.Details["embeddingMode"]);
        Assert.Equal("true", warning.Details["cffCharstringsRetained"]);
        Assert.Equal("false", warning.Details["openTypeLayoutTablesEmbedded"]);
        Assert.Contains("CFF", warning.Details["openTypeTablesEmbedded"], StringComparison.Ordinal);
        Assert.Contains("GSUB", warning.Details["openTypeTablesRemoved"], StringComparison.Ordinal);
        Assert.Contains("GPOS", warning.Details["openTypeLayoutTablesRemoved"], StringComparison.Ordinal);
        Assert.Equal("false", warning.Details["cffCharstringsSubset"]);
        Assert.Equal(warning.Details["glyphCount"], warning.Details["retainedCffGlyphCount"]);
        Assert.True(int.Parse(warning.Details["unusedCffGlyphCount"], CultureInfo.InvariantCulture) > 0);
        Assert.False(string.IsNullOrWhiteSpace(warning.Details["usedGlyphIdsPreview"]));
    }

    private static IReadOnlyList<string> ReadTableTags(byte[] fontData) {
        int tableCount = (fontData[4] << 8) | fontData[5];
        var tags = new List<string>(tableCount);
        for (int index = 0; index < tableCount; index++) {
            int offset = 12 + index * 16;
            tags.Add(Encoding.ASCII.GetString(fontData, offset, 4));
        }

        return tags;
    }

    private static IReadOnlyList<int> ExtractLength1Values(string raw) {
        var values = new List<int>();
        foreach (Match match in Regex.Matches(raw, @"/Length1\s+(\d+)")) {
            values.Add(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture));
        }

        return values;
    }
}
