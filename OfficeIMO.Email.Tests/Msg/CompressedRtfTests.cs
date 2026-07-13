using OfficeIMO.Email;
using System.Threading;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class CompressedRtfTests {
    private static readonly byte[] OfficialSimpleExample = {
        0x2d, 0x00, 0x00, 0x00, 0x2b, 0x00, 0x00, 0x00, 0x4c, 0x5a, 0x46, 0x75, 0xf1, 0xc5, 0xc7, 0xa7,
        0x03, 0x00, 0x0a, 0x00, 0x72, 0x63, 0x70, 0x67, 0x31, 0x32, 0x35, 0x42, 0x32, 0x0a, 0xf3, 0x20,
        0x68, 0x65, 0x6c, 0x09, 0x00, 0x20, 0x62, 0x77, 0x05, 0xb0, 0x6c, 0x64, 0x7d, 0x0a, 0x80, 0x0f,
        0xa0
    };

    private static readonly byte[] OfficialCrossWriteExample = {
        0x1a, 0x00, 0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x4c, 0x5a, 0x46, 0x75, 0xe2, 0xd4, 0x4b, 0x51,
        0x41, 0x00, 0x04, 0x20, 0x57, 0x58, 0x59, 0x5a, 0x0d, 0x6e, 0x7d, 0x01, 0x0e, 0xb0
    };

    [Fact]
    public void DecompressesOfficialMicrosoftExamplesIncludingDictionaryWrap() {
        var diagnostics = new List<EmailDiagnostic>();

        Assert.True(MapiCompressedRtfCodec.TryDecompress(OfficialSimpleExample, 1024, diagnostics,
            "sample-1", CancellationToken.None, out byte[] simple));
        Assert.Equal("{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n", BytePreservingString(simple));

        Assert.True(MapiCompressedRtfCodec.TryDecompress(OfficialCrossWriteExample, 1024, diagnostics,
            "sample-2", CancellationToken.None, out byte[] crossing));
        Assert.Equal("{\\rtf1 WXYZWXYZWXYZWXYZWXYZ}", BytePreservingString(crossing));
        Assert.Empty(diagnostics);
    }

    [Fact]
    public void CompressorRoundTripsLargeRepeatingRtfAndWritesValidCrc() {
        string rtf = string.Concat("{\\rtf1\\ansi ", new string('A', 10000), "}");
        byte[] raw = BytePreservingBytes(rtf);

        byte[] compressed = MapiCompressedRtfCodec.Compress(raw);
        uint storedCrc = MsgBinary.ReadUInt32(compressed, 12);
        uint actualCrc = MapiCompressedRtfCodec.CalculateCrc(compressed, 16, compressed.Length - 16);
        var diagnostics = new List<EmailDiagnostic>();

        Assert.Equal(actualCrc, storedCrc);
        Assert.True(MapiCompressedRtfCodec.TryDecompress(compressed, raw.Length, diagnostics,
            "roundtrip", CancellationToken.None, out byte[] decompressed));
        Assert.Equal(raw, decompressed);
        Assert.Empty(diagnostics);
        Assert.True(compressed.Length < raw.Length);
    }

    [Fact]
    public void CompressorRoundTripsEmptyInputUsingTheSpecifiedSentinelRun() {
        byte[] compressed = MapiCompressedRtfCodec.Compress(Array.Empty<byte>());
        var diagnostics = new List<EmailDiagnostic>();

        Assert.Equal(new byte[] { 0x02, 0x00, 0x0d, 0x00 }, compressed.Skip(16).ToArray());
        Assert.True(MapiCompressedRtfCodec.TryDecompress(compressed, 1, diagnostics,
            "empty", CancellationToken.None, out byte[] decoded));
        Assert.Empty(decoded);
        Assert.Empty(diagnostics);
    }

    [Fact]
    public void ReadsSpecDefinedUncompressedPayloadAndRejectsDamageAndExpansion() {
        const string rtf = "{\\rtf1 uncompressed}";
        byte[] raw = BytePreservingBytes(rtf);
        byte[] uncompressed = new byte[16 + raw.Length];
        MsgBinary.WriteUInt32(uncompressed, 0, checked((uint)(raw.Length + 12)));
        MsgBinary.WriteUInt32(uncompressed, 4, checked((uint)raw.Length));
        MsgBinary.WriteUInt32(uncompressed, 8, 0x414C454D);
        Buffer.BlockCopy(raw, 0, uncompressed, 16, raw.Length);
        var diagnostics = new List<EmailDiagnostic>();

        Assert.True(MapiCompressedRtfCodec.TryDecompress(uncompressed, 1024, diagnostics,
            "uncompressed", CancellationToken.None, out byte[] decoded));
        Assert.Equal(raw, decoded);

        byte[] damaged = (byte[])OfficialSimpleExample.Clone();
        damaged[20] ^= 0x01;
        Assert.False(MapiCompressedRtfCodec.TryDecompress(damaged, 1024, diagnostics,
            "damaged", CancellationToken.None, out _));
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "EMAIL_MSG_RTF_CRC_MISMATCH");

        byte[] oversized = (byte[])OfficialSimpleExample.Clone();
        MsgBinary.WriteUInt32(oversized, 4, 2048);
        Assert.Throws<EmailLimitExceededException>(() => MapiCompressedRtfCodec.TryDecompress(
            oversized, 1024, diagnostics, "oversized", CancellationToken.None, out _));
    }

    [Fact]
    public void MsgRoundTripProjectsRtfAndIsReadableByMsgReaderOracle() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        const string rtf = "{\\rtf1\\ansi\\ansicpg1252\\pard OfficeIMO RTF body\\par}";
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "RTF"
        };
        source.Body.Rtf = rtf;

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(rtf, result.Document.Body.Rtf);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
        MapiProperty compressed = Assert.Single(result.Document.MapiProperties,
            property => property.PropertyId == 0x1009);
        Assert.IsType<byte[]>(compressed.Value);
        Assert.Equal(2, result.Document.MapiProperties.Single(property => property.PropertyId == 0x1016).Value);

        using var stream = new MemoryStream(bytes, writable: false);
        using var oracle = new global::MsgReader.Outlook.Storage.Message(stream, FileAccess.Read, true);
        Assert.Equal(rtf, oracle.BodyRtf);
    }

    [Fact]
    public void MsgReadProjectsOutlookEncapsulatedHtmlThroughOfficeImoRtf() {
        const string rtf = @"{\rtf1\ansi\fromhtml1{\*\htmltag <p><b>Rich</b> message</p>}Plain fallback}";
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "HTML in RTF"
        };
        source.Body.Rtf = rtf;

        EmailReadResult result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg));

        Assert.Equal(rtf, result.Document.Body.Rtf);
        Assert.Equal("<p><b>Rich</b> message</p>", result.Document.Body.Html);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    private static byte[] BytePreservingBytes(string value) {
        byte[] result = new byte[value.Length];
        for (int index = 0; index < value.Length; index++) result[index] = checked((byte)value[index]);
        return result;
    }

    private static string BytePreservingString(byte[] value) {
        char[] result = new char[value.Length];
        for (int index = 0; index < value.Length; index++) result[index] = (char)value[index];
        return new string(result);
    }
}
