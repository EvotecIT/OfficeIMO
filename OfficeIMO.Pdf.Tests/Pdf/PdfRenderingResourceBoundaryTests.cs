using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRenderingResourceBoundaryTests {
    [Fact]
    public void Type1SetCurrentPointDoesNotLeakOperandsIntoTheNextOutlineOperator() {
        byte[] start = { 139, 248, 136, 13, 159, 159, 21 }; // 0 500 hsbw, 20 20 rmoveto
        byte[] finish = { 149, 149, 5, 14 }; // 10 10 rlineto, endchar
        byte[] setCurrentPoint = { 159, 159, 12, 33 };
        PdfType1FontConverter.Result? expected = PdfType1FontConverter.TryConvert(Type1(start.Concat(finish).ToArray()));
        PdfType1FontConverter.Result? actual = PdfType1FontConverter.TryConvert(Type1(start.Concat(setCurrentPoint).Concat(finish).ToArray()));
        Assert.NotNull(expected);
        Assert.NotNull(actual);
        Assert.Equal(expected!.Cff, actual!.Cff);
    }

    [Fact]
    public void Jbig2IntermediateRegionCannotExceedTheCallerByteBudget() {
        // A generic MMR region, much wider than the requested one-pixel output.
        byte[] header = { 0, 0, 0, 1, 38, 0, 1, 0, 0, 0, 18 };
        byte[] region = new byte[18];
        region[3] = 64;
        region[7] = 1;
        region[17] = 1;
        InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
            Jbig2Decoder.Decode(header.Concat(region).ToArray(), null, 1, 1, 32, default));
        Assert.Contains("byte budget", error.Message);
    }

    [Fact]
    public void CyclicJbig2GlobalsFailWithoutRecursingIndefinitely() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("JBIG2Decode");
        dictionary.Items["Width"] = new PdfNumber(1);
        dictionary.Items["Height"] = new PdfNumber(1);
        var parameters = new PdfDictionary();
        dictionary.Items["DecodeParms"] = parameters;
        var stream = new PdfStream(dictionary, new byte[] { 0 });
        parameters.Items["JBIG2Globals"] = stream;
        Assert.False(StreamDecoder.TryDecode(dictionary, stream.Data, 1024, out _));
        Assert.Equal("malformed", StreamDecoder.FindUnsupportedJbig2Feature(stream, null, 1024));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CffGlyphNameBudgetCountsUniqueDecodedNames(bool uniqueNames) {
        // Repeated SIDs share one decoded name. Distinct names must fit the aggregate budget.
        byte[] font = CffWithCustomNames(32, 65536, uniqueNames);
        PdfCffFontProgram? parsed = PdfCffFontProgram.TryParse(font);
        if (uniqueNames) Assert.Null(parsed);
        else {
            Assert.NotNull(parsed);
            Assert.Equal(32, parsed!.GlyphCount);
            Assert.Equal(65536, parsed.GlyphNames![31].Length);
        }
    }

    [Fact]
    public void SubsetTagsDistinguishFontBytesWithTheSameNameAndGlyphs() {
        string first = PdfSubsetTag.Apply("SameName", new[] { 1, 2 }, new byte[] { 1, 2, 3 });
        string second = PdfSubsetTag.Apply("SameName", new[] { 1, 2 }, new byte[] { 1, 2, 4 });
        Assert.NotEqual(first, second);
        Assert.Equal(first, PdfSubsetTag.Apply("SameName", new[] { 1, 2 }, new byte[] { 1, 2, 3 }));
    }

    [Fact]
    public void RegionalIndicatorAnalysisHonorsCancellationAndKeepsPairBoundaries() {
        int[] scalars = Enumerable.Repeat(0x1F1E6, 50000).ToArray();
        var actions = OfficeUnicodeLineBreaker.Analyze(scalars);
        for (int index = 1; index < actions.Length; index++)
            Assert.Equal(index % 2 == 0 ? OfficeUnicodeLineBreaker.BreakAction.Allowed : OfficeUnicodeLineBreaker.BreakAction.Prohibited, actions[index]);
        using var source = new CancellationTokenSource();
        source.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeTextLineBreaks.GetBreakPositions(char.ConvertFromUtf32(0x1F1E6), true, source.Token));
    }

    private static byte[] CffWithCustomNames(int glyphs, int nameLength, bool uniqueNames) {
        byte[] names = Index(new[] { Encoding.ASCII.GetBytes("Bounded") });
        byte[][] custom = Enumerable.Range(0, uniqueNames ? glyphs - 1 : 1)
            .Select(index => Encoding.ASCII.GetBytes(((char)('A' + index % 26)).ToString() + new string('x', nameLength - 1))).ToArray();
        byte[] strings = Index(custom);
        byte[] charStrings = Index(Enumerable.Range(0, glyphs).Select(_ => new byte[] { 14 }).ToArray());
        // Two fixed-width DICT integers: charset offset (15) and CharStrings offset (17).
        byte[] top = new byte[12];
        top[0] = top[6] = 29; top[5] = 15; top[11] = 17;
        int charsetOffset = 4 + names.Length + Index(new[] { top }).Length + strings.Length + 2;
        SetInt(top, 1, charsetOffset);
        SetInt(top, 7, charsetOffset + 1 + (glyphs - 1) * 2);
        using var output = new MemoryStream();
        void Write(byte[] bytes) => output.Write(bytes, 0, bytes.Length);
        Write(new byte[] { 1, 0, 4, 4 }); Write(names); Write(Index(new[] { top })); Write(strings); Write(new byte[2]);
        output.WriteByte(0);
        for (int glyph = 1; glyph < glyphs; glyph++) {
            int sid = 391 + (uniqueNames ? glyph - 1 : 0);
            output.WriteByte((byte)(sid >> 8)); output.WriteByte((byte)sid);
        }
        Write(charStrings);
        return output.ToArray();
    }

    private static byte[] Type1(byte[] glyph, string encoding = "", string glyphName = "A", byte[]? cipherPrefix = null) {
        using var privatePart = new MemoryStream();
        void Write(string text) { byte[] bytes = Encoding.ASCII.GetBytes(text); privatePart.Write(bytes, 0, bytes.Length); }
        privatePart.Write(new byte[4], 0, 4);
        Write("/lenIV -1 def /CharStrings 2 dict dup begin /.notdef 1 RD ");
        privatePart.WriteByte(14);
        Write(" ND /" + glyphName + " " + glyph.Length + " RD ");
        privatePart.Write(glyph, 0, glyph.Length);
        Write(" ND end");
        byte[] bytes = privatePart.ToArray();
        ushort state = 55665;
        for (int index = 0; index < bytes.Length; index++) {
            if (cipherPrefix != null && index < 4) bytes[index] = cipherPrefix[index];
            else bytes[index] ^= (byte)(state >> 8);
            state = unchecked((ushort)((bytes[index] + state) * 52845 + 22719));
        }
        return Encoding.ASCII.GetBytes("%!PS-AdobeFont-1.0\n/FontName /Review def\n" + encoding + "\ncurrentfile eexec\n" + BitConverter.ToString(bytes).Replace("-", string.Empty));
    }

    private static byte[] Index(byte[][] values) {
        using var output = new MemoryStream();
        output.WriteByte((byte)(values.Length >> 8)); output.WriteByte((byte)values.Length); output.WriteByte(4);
        int offset = 1;
        foreach (byte[] value in values) { WriteInt(output, offset); offset += value.Length; }
        WriteInt(output, offset);
        foreach (byte[] value in values) output.Write(value, 0, value.Length);
        return output.ToArray();
    }

    private static void SetInt(byte[] bytes, int offset, int value) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (24 - index * 8));
    }

    private static void WriteInt(Stream stream, int value) {
        for (int index = 0; index < 4; index++) stream.WriteByte((byte)(value >> (24 - index * 8)));
    }
}
