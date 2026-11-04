using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRenderingResourceBoundaryTests {
    [Theory]
    [InlineData(1023)]
    [InlineData(1024)]
    [InlineData(8188)]
    [InlineData(8189)]
    public void CompactSubsetKeepsSparseBmpMappingsAcrossCmapLimits(int count) {
        int[] scalars = Enumerable.Range(0, count).Select(index => 32 + 2 * index).ToArray();
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFont(scalars));
        program.RecordGlyphUsage(1, scalars[0]);
        OfficeTrueTypeFont? subset = OfficeTrueTypeFont.TryLoad(program.BuildSubsetFontFile());
        Assert.NotNull(subset);
        foreach (int scalar in scalars) {
            Assert.True(subset!.TryGetGlyphMetrics(scalar, out int glyph, out _), $"Missing U+{scalar:X4}");
            Assert.Equal(1, glyph);
        }
    }

    [Theory]
    [InlineData("ISOLatin1Encoding", 193, "Aacute", "\u00C1")]
    [InlineData("StandardEncoding 256 array copy dup 65 /B put", 65, "B", "B")]
    [InlineData("StandardEncoding 256 array copy dup\t65 /B put", 65, "B", "B")]
    [InlineData("StandardEncoding 256 array copy % dup 65 /B put\nreadonly", 65, "A", "A")]
    public void Type1BuiltinEncodingAgreesWithExtractedText(string encoding, int code, string glyphName, string text) {
        byte[] data = Type1(new byte[] { 139, 248, 136, 13, 159, 159, 21, 149, 149, 5, 14 },
            "/Encoding " + encoding + " def", glyphName);
        PdfType1FontConverter.Result? converted = PdfType1FontConverter.TryConvert(data);
        Assert.NotNull(converted);
        Assert.Equal(glyphName, converted!.BuiltinEncoding[code]);
        PdfDictionary font = Type1Resource(data);
        var graph = new Dictionary<int, PdfIndirectObject>();
        PdfFontResource resource = ResourceResolver.CreateFontResource("F", font, graph);
        Assert.Null(resource.RetainedDrawingResource);
        Assert.Equal(text, ResourceResolver.CreateBudgetedDecoder(resource)(new[] { (byte)code }, 8));
        Assert.Equal(text, ResourceResolver.CreateBudgetedDecoder(resource, true, false)(new[] { (byte)code }, 8));
        Assert.Null(resource.RetainedDrawingResource); // text metadata never converts outlines
        Assert.Equal(PdfTextDecodingSource.FontEncoding, ResourceResolver.GetDecodingEvidence(resource, new[] { (byte)code }, null).Source);
        Assert.Equal(PdfTextDecodingSource.FontEncoding, ResourceResolver.GetDecodingEvidence(resource, new[] { (byte)code }, null, true, false).Source);
        PdfFontResource drawing = resource.ForDrawing();
        Assert.NotNull(drawing.DrawingProgram);
        Assert.True(drawing.DrawingProgram!.GlyphForCode(code) > 0);

        var differences = new PdfDictionary();
        differences.Items["Differences"] = new PdfArray { Items = { new PdfNumber(code), new PdfName("A") } };
        font.Items["Encoding"] = differences;
        Assert.Equal("A", ResourceResolver.CreateBudgetedDecoder(ResourceResolver.CreateFontResource("F", font, graph))(new[] { (byte)code }, 8));
        font.Items["Encoding"] = new PdfName("WinAnsiEncoding");
        Assert.Equal(PdfWinAnsiEncoding.Decode(new[] { (byte)code }, 8),
            ResourceResolver.CreateBudgetedDecoder(ResourceResolver.CreateFontResource("F", font, graph))(new[] { (byte)code }, 8));
    }

    [Theory]
    [InlineData("BogusEncoding")]
    [InlineData("StandardEncodingBogus")]
    public void UnknownType1EncodingDoesNotClaimAnAuthoritativeEmptyVector(string name) {
        byte[] data = Type1(new byte[] { 14 }, "/Encoding " + name + " def");
        Assert.Null(PdfType1FontConverter.TryConvert(data));
        Assert.Null(PdfType1FontConverter.TryReadBuiltinEncoding(data));
    }

    [Fact]
    public void OversizedType1EncodingNameIsRejectedBeforeTextDecoding() {
        byte[] data = Type1(new byte[] { 14 },
            "/Encoding 256 array dup 65 /" + new string('Z', 4096) + " put def");

        Assert.Null(PdfType1FontConverter.TryReadBuiltinEncoding(data));
    }

    [Fact]
    public void RepeatedUnknownType1EncodingNameUsesTheFallbackConsistently() {
        byte[] data = Type1(new byte[] { 14 },
            "/Encoding 256 array dup 65 /" + new string('Z', 4095) + " put def");
        PdfFontResource resource = ResourceResolver.CreateFontResource("F", Type1Resource(data),
            new Dictionary<int, PdfIndirectObject>());
        byte[] repeated = Enumerable.Repeat((byte)65, 20_000).ToArray();

        Assert.Equal(new string('A', repeated.Length), ResourceResolver.CreateBudgetedDecoder(resource)(repeated, repeated.Length));
        Assert.Equal(PdfTextDecodingSource.WinAnsiFallback,
            ResourceResolver.GetDecodingEvidence(resource, repeated, null).Source);
        Assert.Throws<PdfUnsupportedTextMappingException>(() =>
            ResourceResolver.CreateBudgetedDecoder(resource, preferToUnicode: true, useWinAnsiFallback: false)(repeated, repeated.Length));
    }

    [Fact]
    public void RepeatedType1TextRunsShareMappingsAcrossDifferencesAndPolicies() {
        byte[] data = Type1(new byte[] { 14 },
            "/Encoding 256 array dup 65 /" + new string('Z', 4095) + " put def");
        PdfDictionary font = Type1Resource(data);
        var differences = new PdfDictionary();
        differences.Items["Differences"] = new PdfArray { Items = { new PdfNumber(66), new PdfName("B") } };
        font.Items["Encoding"] = differences;
        var baseFonts = new PdfDictionary(); baseFonts.Items["F"] = font;
        var aliasFonts = new PdfDictionary(); aliasFonts.Items["Alias"] = font;
        var baseResources = new PdfDictionary(); baseResources.Items["Font"] = baseFonts;
        var aliasResources = new PdfDictionary(); aliasResources.Items["Font"] = aliasFonts;
        var graph = new Dictionary<int, PdfIndirectObject>();
        var cache = new PdfFontResourceCache();
        PdfFontResource resource = cache.GetOrCreate(baseResources, graph).Fonts["F"];
        Func<byte[], int, string> defaultDecoder = ResourceResolver.CreateBudgetedDecoder(resource);
        byte[] run = { 65, 66 };

        for (int index = 0; index < 32; index++) {
            PdfFontResource alias = cache.GetOrCreate(aliasResources, graph).Fonts["Alias"];
            Assert.Same(resource.BuiltinEncodingSource, alias.BuiltinEncodingSource);
            Func<byte[], int, string> strictDecoder = ResourceResolver.CreateBudgetedDecoder(alias,
                preferToUnicode: true, useWinAnsiFallback: false);
            Assert.Equal("AB", defaultDecoder(run, 2));
            Assert.Equal(PdfTextDecodingSource.WinAnsiFallback,
                ResourceResolver.GetDecodingEvidence(alias, run, null).Source);
            Assert.Equal(PdfTextDecodingSource.Unmappable,
                ResourceResolver.GetDecodingEvidence(alias, run, null,
                    preferToUnicode: true, useWinAnsiFallback: false).Source);
            Assert.Throws<PdfUnsupportedTextMappingException>(() => strictDecoder(run, 2));
        }
    }

    [Theory]
    [InlineData(".notdef")]
    [InlineData("missingglyph")]
    public void ExplicitUnmappedDifferenceCannotRegainItsBuiltinGlyph(string glyphName) {
        byte[] data = Type1(new byte[] { 139, 248, 136, 13, 159, 159, 21, 149, 149, 5, 14 }, "/Encoding StandardEncoding def");
        PdfDictionary font = Type1Resource(data);
        var encoding = new PdfDictionary();
        encoding.Items["Differences"] = new PdfArray { Items = { new PdfNumber(65), new PdfName(glyphName) } };
        font.Items["Encoding"] = encoding;
        PdfFontResource resource = ResourceResolver.CreateFontResource("F", font, new Dictionary<int, PdfIndirectObject>());
        Assert.Equal(string.Empty, ResourceResolver.CreateBudgetedDecoder(resource)(new byte[] { 65 }, 8));
        Assert.Throws<PdfUnsupportedTextMappingException>(() => ResourceResolver.CreateBudgetedDecoder(resource, true, false)(new byte[] { 65 }, 8));
        PdfFontResource drawing = resource.ForDrawing();
        Assert.NotNull(drawing.DrawingProgram);
        Assert.Equal(0, drawing.DrawingProgram!.GlyphForCode(65));
    }

    [Fact]
    public void Type1EncodingNamesInCommentsAndStringsAreInert() {
        byte[] data = Type1(new byte[] { 14 }, "% /Encoding BogusEncoding def\n(/Encoding BogusEncoding def) pop\n/Encoding StandardEncoding def");
        Assert.Equal("A", PdfType1FontConverter.TryReadBuiltinEncoding(data)![65]);
    }

    [Fact]
    public void DeferredDrawingAliasesShareConversionAndCancellationDoesNotPoisonIt() {
        byte[] data = Type1(new byte[] { 139, 248, 136, 13, 159, 159, 21, 149, 149, 5, 14 }, "/Encoding StandardEncoding def");
        PdfDictionary font = Type1Resource(data);
        var fonts = new PdfDictionary(); fonts.Items["F1"] = font; fonts.Items["F2"] = font;
        var resources = new PdfDictionary(); resources.Items["Font"] = fonts;
        var graph = new Dictionary<int, PdfIndirectObject>();
        var cache = new PdfFontResourceCache();
        PdfFontResourceSet metadata = cache.GetOrCreate(resources, graph);
        Assert.True(metadata.Fonts["F1"].HasDeferredDrawing);
        Assert.Null(metadata.Fonts["F1"].RetainedDrawingResource);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => metadata.Fonts["F2"].ForDrawing(cancelled.Token));
        Assert.Null(metadata.Fonts["F1"].RetainedDrawingResource);
        PdfFontResourceSet drawing = cache.GetOrCreate(resources, graph, forDrawing: true);
        Assert.NotNull(drawing.Fonts["F1"].DrawingProgram);
        Assert.Same(drawing.Fonts["F1"].DrawingProgram, drawing.Fonts["F2"].DrawingProgram);
        Assert.Same(drawing.Fonts["F1"].DrawingProgram, cache.GetFonts(resources, graph, forDrawing: true)["F1"].DrawingProgram);
        Assert.Equal("F2", drawing.Fonts["F2"].ResourceName);
    }

    private static PdfDictionary Type1Resource(byte[] program) {
        var descriptor = new PdfDictionary(); descriptor.Items["FontFile"] = new PdfStream(new PdfDictionary(), program);
        var font = new PdfDictionary(); font.Items["Subtype"] = new PdfName("Type1");
        font.Items["BaseFont"] = new PdfName("Review"); font.Items["FontDescriptor"] = descriptor;
        return font;
    }
}
