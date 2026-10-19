using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfRenderingResourceBoundaryTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PfbBinarySegmentIsNotGuessedFromItsFirstCiphertextBytes(bool hexDigits) {
        byte[] prefix = hexDigits ? Encoding.ASCII.GetBytes("ABCD") : new byte[] { 32, 10, 9, 13 };
        byte[] pfa = Type1(new byte[] { 139, 248, 136, 13, 159, 159, 21, 149, 149, 5, 14 },
            "/Encoding StandardEncoding def", cipherPrefix: prefix);
        string text = Encoding.ASCII.GetString(pfa);
        int start = text.IndexOf("eexec", StringComparison.Ordinal) + 6;
        byte[] clear = Encoding.ASCII.GetBytes(text.Substring(0, start));
        string hex = text.Substring(start);
        byte[] binary = Enumerable.Range(0, hex.Length / 2).Select(index => Convert.ToByte(hex.Substring(index * 2, 2), 16)).ToArray();
        using var pfb = new MemoryStream();
        void Segment(int type, byte[] data) {
            pfb.WriteByte(128); pfb.WriteByte((byte)type);
            for (int index = 0; index < 4; index++) pfb.WriteByte((byte)(data.Length >> (index * 8)));
            pfb.Write(data, 0, data.Length);
        }
        Segment(1, clear); Segment(2, binary); pfb.WriteByte(128); pfb.WriteByte(3);
        PdfType1FontConverter.Result? expected = PdfType1FontConverter.TryConvert(pfa);
        PdfType1FontConverter.Result? actual = PdfType1FontConverter.TryConvert(pfb.ToArray());
        Assert.NotNull(expected); Assert.NotNull(actual);
        Assert.Equal(expected!.Cff, actual!.Cff);
        Assert.Equal("A", actual.BuiltinEncoding[65]);
    }

    [Theory]
    [InlineData("0 .001 -.001 0 0 0")]
    [InlineData(".001 .0002 0 .001 0 0")]
    [InlineData(".001 0 0 .002 0 0")]
    [InlineData(".001 0 0 .001 .1 0")]
    [InlineData("-.001 0 0 .001 0 0")]
    public void AffineCffMatricesDoNotSilentlyLoseTheirAuthoredTerms(string matrix) {
        byte[] data = Type1(new byte[] { 139, 248, 136, 13, 159, 159, 21, 149, 149, 5, 14 },
            "/FontMatrix [" + matrix + "] def /Encoding StandardEncoding def");
        PdfFontResource font = ResourceResolver.CreateFontResource("F", Type1Resource(data), new Dictionary<int, PdfIndirectObject>());
        Assert.Equal("Type1", font.EmbeddedProgramSubtype);
        Assert.Null(font.ForDrawing().DrawingProgram);
    }
    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void SelectedCidFontDictionariesMustShareTheSupportedMatrix(bool affineSecond, bool supported) {
        byte[] cff = CffWithTwoFontMatrices(affineSecond);
        PdfCffFontProgram? parsed = PdfCffFontProgram.TryParse(cff);
        Assert.NotNull(parsed);
        Assert.Equal(supported, parsed!.HasUniformFontMatrix);
        var descriptor = new PdfDictionary();
        var streamDictionary = new PdfDictionary(); streamDictionary.Items["Subtype"] = new PdfName("CIDFontType0C");
        descriptor.Items["FontFile3"] = new PdfStream(streamDictionary, cff);
        var descendant = new PdfDictionary(); descendant.Items["Subtype"] = new PdfName("CIDFontType0");
        descendant.Items["BaseFont"] = new PdfName("Review"); descendant.Items["FontDescriptor"] = descriptor;
        var descendants = new PdfArray(); descendants.Items.Add(descendant);
        var font = new PdfDictionary(); font.Items["Subtype"] = new PdfName("Type0");
        font.Items["BaseFont"] = new PdfName("Review"); font.Items["Encoding"] = new PdfName("Identity-H");
        font.Items["DescendantFonts"] = descendants;
        PdfFontResource resource = ResourceResolver.CreateFontResource("F", font, new Dictionary<int, PdfIndirectObject>());
        Assert.Equal(supported, resource.ForDrawing().DrawingProgram != null);
    }

    private static byte[] CffWithTwoFontMatrices(bool affineSecond) {
        byte[] names = Index(new[] { Encoding.ASCII.GetBytes("TwoMatrices") });
        // ROS, then four fixed-width offsets: charset, CharStrings, FDArray and FDSelect.
        byte[] ros = { 139, 139, 139, 12, 30 };
        // The top matrix is .001 times identity, concatenated with each FD matrix.
        byte[] matrix = { 30, 0x0A, 0x00, 0x1F, 139, 139, 30, 0x0A, 0x00, 0x1F, 139, 139, 12, 7 };
        byte[] top = ros.Concat(matrix).Concat(new byte[26]).ToArray();
        int entries = ros.Length + matrix.Length;
        int position = 4 + names.Length + Index(new[] { top }).Length + 4;
        byte[] charset = { 0, 0, 1 };
        byte[] charStrings = Index(new[] { new byte[] { 14 }, new byte[] { 139, 139, 21, 149, 149, 5, 14 } });
        byte[] fd0 = { 140, 139, 139, 140, 139, 139, 12, 7 };
        byte[] fd1 = { 140, 139, (byte)(affineSecond ? 140 : 139), 140, 139, 139, 12, 7 };
        byte[] dictionaries = Index(new[] { fd0, fd1 });
        void Offset(int at, int value, byte op, bool escaped = false) {
            top[at] = 29; SetInt(top, at + 1, value);
            if (escaped) { top[at + 5] = 12; top[at + 6] = op; } else top[at + 5] = op;
        }
        Offset(entries, position, 15); position += charset.Length;
        Offset(entries + 6, position, 17); position += charStrings.Length;
        Offset(entries + 12, position, 36, true); position += dictionaries.Length;
        Offset(entries + 19, position, 37, true);
        return new byte[] { 1, 0, 4, 4 }.Concat(names).Concat(Index(new[] { top })).Concat(new byte[4])
            .Concat(charset).Concat(charStrings).Concat(dictionaries).Concat(new byte[] { 0, 0, 1 }).ToArray();
    }

}
