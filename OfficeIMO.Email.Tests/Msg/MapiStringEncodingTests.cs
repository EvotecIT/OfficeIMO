using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MapiStringEncodingTests {
    static MapiStringEncodingTests() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Fact]
    public void ResolvesDbcsInternetCodePageForString8Content() {
        byte[] propertyStream = BuildPropertyStream(0x3FDE, 932);
        MapiStringEncodingContext context = MapiStringEncodingContext.Resolve(propertyStream, 32, null);
        byte[] bytes = Encoding.GetEncoding(932).GetBytes("日本語の添付.msg");
        var diagnostics = new List<EmailDiagnostic>();

        string decoded = context.Decode(bytes, diagnostics, "subject");

        Assert.Equal(932, context.PrimaryCodePage);
        Assert.Equal("日本語の添付.msg", decoded);
        Assert.Empty(diagnostics);
    }

    [Fact]
    public void RejectsUtf16AsString8CodePageAndUsesLocaleAnsiPage() {
        byte[] propertyStream = new byte[64];
        WriteProperty(propertyStream, 32, 0x3FFD, 1200);
        WriteProperty(propertyStream, 48, 0x3FF1, 1045); // pl-PL -> Windows-1250
        MapiStringEncodingContext context = MapiStringEncodingContext.Resolve(propertyStream, 32, null);
        byte[] bytes = Encoding.GetEncoding(1250).GetBytes("Zażółć gęślą");

        string decoded = context.Decode(bytes, new List<EmailDiagnostic>(), "body");

        Assert.Equal(1250, context.PrimaryCodePage);
        Assert.Equal("Zażółć gęślą", decoded);
    }

    private static byte[] BuildPropertyStream(ushort propertyId, int value) {
        byte[] stream = new byte[48];
        WriteProperty(stream, 32, propertyId, value);
        return stream;
    }

    private static void WriteProperty(byte[] stream, int offset, ushort propertyId, int value) {
        MsgBinary.WriteUInt32(stream, offset, ((uint)propertyId << 16) | (ushort)MapiPropertyType.Integer32);
        MsgBinary.WriteUInt32(stream, offset + 8, unchecked((uint)value));
    }
}
