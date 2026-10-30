using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOpenTypeCffCompactEmbeddingTests {
    [Fact]
    public void CidSubsetIgnoresUnusedFontDictionaryPrivateRanges() {
        static byte[] Index(params byte[][] objects) {
            if (objects.Length == 0) return new byte[] { 0, 0 };
            var bytes = new List<byte> { (byte)(objects.Length >> 8), (byte)objects.Length, 1 };
            int offset = 1;
            foreach (byte[] value in objects) {
                bytes.Add(checked((byte)offset));
                offset += value.Length;
            }
            bytes.Add(checked((byte)offset));
            foreach (byte[] value in objects) bytes.AddRange(value);
            return bytes.ToArray();
        }

        static byte[] TopDictionary(int charStrings, int fdArray, int fdSelect) {
            var bytes = new List<byte> { 139, 139, 139, 12, 30 };
            void AddOffset(int value, params byte[] op) {
                bytes.Add(29);
                bytes.Add((byte)(value >> 24));
                bytes.Add((byte)(value >> 16));
                bytes.Add((byte)(value >> 8));
                bytes.Add((byte)value);
                bytes.AddRange(op);
            }
            AddOffset(charStrings, 17);
            AddOffset(fdArray, 12, 36);
            AddOffset(fdSelect, 12, 37);
            return bytes.ToArray();
        }

        byte[] names = Index(Encoding.ASCII.GetBytes("CID proof"));
        byte[] strings = Index();
        byte[] globalSubrs = Index();
        byte[] charStrings = Index(new byte[] { 14 }, new byte[] { 14 }, new byte[] { 14 });
        byte[] fdSelect = { 0, 0, 0, 1 }; // glyphs 0 and 1 use FD 0; glyph 2 uses FD 1.
        byte[] fdArray = Index(Array.Empty<byte>(), new byte[] { 139, 29, 0x7f, 0xff, 0xff, 0xff, 18 });
        byte[] topPlaceholder = Index(TopDictionary(0, 0, 0));
        int charStringsOffset = 4 + names.Length + topPlaceholder.Length + strings.Length + globalSubrs.Length;
        int fdSelectOffset = charStringsOffset + charStrings.Length;
        int fdArrayOffset = fdSelectOffset + fdSelect.Length;
        byte[] top = Index(TopDictionary(charStringsOffset, fdArrayOffset, fdSelectOffset));
        Assert.Equal(topPlaceholder.Length, top.Length);
        byte[] cff = new byte[] { 1, 0, 4, 4 }
            .Concat(names).Concat(top).Concat(strings).Concat(globalSubrs)
            .Concat(charStrings).Concat(fdSelect).Concat(fdArray).ToArray();

        PdfCffCharStringSubset? subset = PdfCffCidSubsetter.TryCreate(cff, new[] { 1 }, 3);
        Assert.NotNull(subset);
        Assert.True(subset!.IsSubset);
        Assert.Equal(2, subset.RetainedGlyphCount);
    }

    [Fact]
    public void PdfOpenTypeCffFontProgram_BuildsCompactOpenTypeFontFileWithoutLayoutTables() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        byte[] full = File.ReadAllBytes(fontPath!);
        PdfOpenTypeCffFontProgram program = PdfOpenTypeCffFontProgram.Parse(full, "OfficeIMO Source Serif CFF");
        program.EncodeTextAsGlyphHex("Subset CFF Łódź");
        PdfOpenTypeCffCompactFontFile plan = program.BuildCompactOpenTypeFontFilePlan();
        byte[] compact = plan.Data;
        PdfOpenTypeFontInfo compactInfo = PdfOpenTypeFontInspector.Inspect(compact, "OfficeIMO Source Serif CFF");
        IReadOnlyList<string> compactTables = ReadTableTags(compact);
        PdfCffCharStringSubset subset = Assert.IsType<PdfCffCharStringSubset>(plan.CharStringSubset);

        Assert.True(compact.Length < full.Length);
        Assert.True(subset.IsSubset);
        Assert.True(subset.PrunedGlyphCount > 1000);
        Assert.True(subset.RetainedGlyphCount > 1);
        Assert.Equal(compactInfo.GlyphCount, subset.GlyphCount);
        Assert.True(subset.SubsetProgramBytes < subset.OriginalProgramBytes);
        Assert.True(subset.SubsetIndexBytes < subset.OriginalIndexBytes);
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
        var options = new PdfCreateOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Source Serif CFF");

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Compact CFF Łódź"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        // Used glyphs only, as a CID-keyed bare CFF program whose CIDs are the source glyph ids.
        Match fontFile = Assert.Single(Regex.Matches(raw, @"<< /Length (\d+) /Subtype /CIDFontType0C >>\s*stream\r?\n").Cast<Match>());
        int embeddedLength = int.Parse(fontFile.Groups[1].Value, CultureInfo.InvariantCulture);
        byte[] cff = bytes.Skip(fontFile.Index + fontFile.Length).Take(embeddedLength).ToArray();

        Assert.Contains("Compact CFF Łódź", extracted, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", raw, StringComparison.Ordinal);
        Assert.Equal(1, cff[0]);
        Assert.InRange(embeddedLength, 1, fontData.Length / 4);
        // The inspector executes every charstring of the CID-keyed program, including subroutine calls.
        PdfPrintProductionStructureEvidence evidence = PdfReadDocument.Open(bytes).InspectPrintProductionStructure();
        Assert.Equal(1, evidence.FontResourceCount);
        Assert.Equal(0, evidence.UnembeddedFontResourceCount);
        Assert.Equal(0, evidence.UninspectableFontResourceCount);
        Assert.DoesNotContain(report.Warnings, item => item.Code == "opentype-cff-charstrings-not-subset");
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
}
