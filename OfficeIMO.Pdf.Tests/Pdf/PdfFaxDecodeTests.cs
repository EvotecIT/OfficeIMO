using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFaxDecodeTests {
    [Theory]
    [InlineData("group3", 0, false)]
    [InlineData("group4", -1, false)]
    [InlineData("group3-options1", 2, false)]
    [InlineData("group3-options4", 0, true)]
    [InlineData("group3-options5", 2, true)]
    public void Fax_DecodesIndependentTiffStripsIncludingOddRowsAndLongRuns(string encoding, int k, bool aligned) {
        string root = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "Scans");
        byte[] encoded = File.ReadAllBytes(Path.Combine(root, "fax-pattern." + encoding));
        byte[] expected = File.ReadAllBytes(Path.Combine(root, "fax-pattern.pixels"));
        PdfDictionary dictionary = FaxDictionary(3001, 17, k, blackIsOne: true);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EndOfLine"] = new PdfBoolean(k >= 0);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EncodedByteAlign"] = new PdfBoolean(aligned);
        byte[] decoded = StreamDecoder.DecodeRequired(dictionary, encoded);
        Assert.Equal(expected, decoded);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["BlackIs1"] = new PdfBoolean(false);
        Assert.Equal(expected.Select(value => (byte)~value), StreamDecoder.DecodeRequired(dictionary, encoded));
        Assert.Throws<PdfReadLimitException>(() => StreamDecoder.DecodeRequired(dictionary, encoded, maxOutputBytes: expected.Length - 1));
    }

    [Theory]
    [InlineData(0, "00110101 000101")] // White zero, black eight.
    [InlineData(-1, "001 00110101 000101")]
    [InlineData(2, "000000000001 1 00110101 000101")]
    public void Fax_DecodesOneDimensionalHorizontalAndMixedModes(int k, string bits) {
        Assert.Equal(new byte[] { 255 }, StreamDecoder.DecodeRequired(FaxDictionary(8, 1, k, true), Pack(bits)));
    }

    [Fact]
    public void Fax_RejectsTruncationInvalidRunsAndMissingEndMarker() {
        PdfDictionary dictionary = FaxDictionary(8, 1, 0, true);
        Assert.Throws<InvalidDataException>(() => StreamDecoder.DecodeRequired(dictionary, Array.Empty<byte>()));
        Assert.Throws<InvalidDataException>(() => StreamDecoder.DecodeRequired(dictionary, Pack("10100"))); // White nine.
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EndOfBlock"] = new PdfBoolean(true);
        Assert.Throws<InvalidDataException>(() => StreamDecoder.DecodeRequired(dictionary, Pack("10011"))); // White eight, no RTC.
        Assert.ThrowsAny<OperationCanceledException>(() => StreamDecoder.DecodeRequired(dictionary, Pack("10011"),
            cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void Fax_UsesImageHeightAndValidatesDeclaredWidthAndEndMarker() {
        PdfDictionary dictionary = FaxDictionary(8, 0, -1, true);
        dictionary.Items["Width"] = new PdfNumber(8);
        dictionary.Items["Height"] = new PdfNumber(1);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EndOfBlock"] = new PdfBoolean(true);
        byte[] encoded = Pack("1 000000000001 000000000001");
        Assert.Equal(new byte[] { 0 }, StreamDecoder.DecodeRequired(dictionary, encoded));
        dictionary.Items["Width"] = new PdfNumber(7);
        Assert.Throws<InvalidDataException>(() => StreamDecoder.DecodeRequired(dictionary, encoded));
    }

    [Fact]
    public void Fax_HonorsByteAlignmentAndDecodeFilterPrefixes() {
        PdfDictionary dictionary = FaxDictionary(8, 2, 0, true);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EncodedByteAlign"] = new PdfBoolean(true);
        byte[] data = Pack("10011 000 00110101 000101 00");
        Assert.Equal(new byte[] { 0, 255 }, StreamDecoder.DecodeRequired(dictionary, data));
        var filters = new PdfArray();
        filters.Items.Add(new PdfName("ASCIIHexDecode"));
        filters.Items.Add(new PdfName("CCF"));
        var parameters = new PdfArray();
        parameters.Items.Add(PdfNull.Instance);
        parameters.Items.Add(dictionary.Items["DecodeParms"]);
        dictionary.Items["Filter"] = filters;
        dictionary.Items["DecodeParms"] = parameters;
        byte[] hex = Encoding.ASCII.GetBytes(BitConverter.ToString(data).Replace("-", "") + ">");
        Assert.Equal(new byte[] { 0, 255 }, StreamDecoder.DecodeRequired(dictionary, hex));
    }

    [Theory]
    [InlineData(-1, "1 1", 2)]
    [InlineData(0, "10011 10011", 6)]
    [InlineData(2, "1 10011 1 10011", 6)]
    public void Fax_EndOfBlockOverridesConflictingRowHintWithinImageHeight(int k, string rows, int markers) {
        PdfDictionary dictionary = FaxDictionary(8, 1, k, true);
        dictionary.Items["Height"] = new PdfNumber(2);
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["EndOfBlock"] = new PdfBoolean(true);
        string end = string.Concat(Enumerable.Repeat("000000000001" + (k > 0 ? "1" : ""), markers));
        byte[] encoded = Pack(rows + end);
        Assert.Equal(new byte[] { 0, 0 }, StreamDecoder.DecodeRequired(dictionary, encoded));
        ((PdfDictionary)dictionary.Items["DecodeParms"]).Items["Rows"] = new PdfNumber(3);
        Assert.Equal(new byte[] { 0, 0 }, StreamDecoder.DecodeRequired(dictionary, encoded));
        dictionary.Items["Height"] = new PdfNumber(1);
        Assert.Throws<InvalidDataException>(() => StreamDecoder.DecodeRequired(dictionary, encoded));
    }

    private static PdfDictionary FaxDictionary(int columns, int rows, int k, bool blackIsOne) {
        var parameters = new PdfDictionary();
        parameters.Items["Columns"] = new PdfNumber(columns);
        parameters.Items["Rows"] = new PdfNumber(rows);
        parameters.Items["K"] = new PdfNumber(k);
        parameters.Items["BlackIs1"] = new PdfBoolean(blackIsOne);
        parameters.Items["EndOfBlock"] = new PdfBoolean(false);
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("CCITTFaxDecode");
        dictionary.Items["DecodeParms"] = parameters;
        return dictionary;
    }

    internal static byte[] Pack(string codes) {
        string bits = codes.Replace(" ", "");
        var bytes = new byte[(bits.Length + 7) / 8];
        for (int i = 0; i < bits.Length; i++) if (bits[i] == '1') bytes[i / 8] |= (byte)(128 >> (i & 7));
        return bytes;
    }
}
