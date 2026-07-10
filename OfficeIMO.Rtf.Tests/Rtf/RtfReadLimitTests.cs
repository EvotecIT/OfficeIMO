using OfficeIMO.Rtf;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfReadLimitTests {
    [Fact]
    public void Untrusted_Profile_Is_Bounded_And_Blocks_External_Object_References() {
        RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();

        Assert.Equal(128, options.MaxDepth);
        Assert.Equal(16L * 1024 * 1024, options.MaxInputBytes);
        Assert.Equal(1_000_000, options.MaxTokenCount);
        Assert.Equal(4 * 1024 * 1024, options.MaxBinaryBytesPerPayload);
        Assert.Equal(256, options.MaxImageCount);
        Assert.Equal(32, options.MaxObjectCount);
        Assert.Equal(100_000, options.MaxSemanticBlockCount);
        Assert.False(options.ReadEmbeddedObjects);
        Assert.False(options.ReadFileReferences);
        Assert.Equal(RtfHyperlinkReadPolicy.WebAndMailOnly, options.HyperlinkPolicy);
    }

    [Theory]
    [InlineData("RtfInputCharacterLimitExceeded", "MaxInputCharacters")]
    [InlineData("RtfTokenLimitExceeded", "MaxTokenCount")]
    [InlineData("RtfGroupLimitExceeded", "MaxGroupCount")]
    [InlineData("RtfTextCharacterLimitExceeded", "MaxTextCharacters")]
    public void Syntax_Limits_Throw_Stable_Exceptions(string expectedCode, string limitSource) {
        const string rtf = @"{\rtf1{nested}Visible text}";
        var options = new RtfReadOptions();
        switch (limitSource) {
            case "MaxInputCharacters": options.MaxInputCharacters = 5; break;
            case "MaxTokenCount": options.MaxTokenCount = 2; break;
            case "MaxGroupCount": options.MaxGroupCount = 1; break;
            case "MaxTextCharacters": options.MaxTextCharacters = 3; break;
        }

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Read(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
        Assert.Equal(limitSource, exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void Byte_Limit_Is_Checked_Before_Seekable_Stream_Is_Read() {
        byte[] bytes = Encoding.ASCII.GetBytes(@"{\rtf1 Too large}");
        using var stream = new MemoryStream(bytes);
        var options = new RtfReadOptions { MaxInputBytes = 4 };

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Load(stream, options));

        Assert.Equal("RtfInputByteLimitExceeded", exception.Code);
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public async Task Byte_Limit_Is_Enforced_By_Async_Stream_Load() {
        byte[] bytes = Encoding.ASCII.GetBytes(@"{\rtf1 Too large}");
        using var stream = new MemoryStream(bytes);
        var options = new RtfReadOptions { MaxInputBytes = 4 };

        RtfReadLimitException exception = await Assert.ThrowsAsync<RtfReadLimitException>(() =>
            RtfDocument.LoadAsync(stream, options));

        Assert.Equal("RtfInputByteLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadOptions.MaxInputBytes), exception.LimitSource);
    }

    [Theory]
    [InlineData("RtfBinaryPayloadLimitExceeded", 3, null)]
    [InlineData("RtfTotalBinaryLimitExceeded", null, 3)]
    public void Binary_Limits_Are_Checked_Before_Payload_Allocation(string expectedCode, int? perPayload, int? total) {
        const string rtf = "{\\rtf1\\bin4 abcd}";
        var options = new RtfReadOptions {
            MaxBinaryBytesPerPayload = perPayload,
            MaxTotalBinaryBytes = total
        };

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Read(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
        Assert.Equal(4, exception.Actual);
        Assert.Equal(3, exception.Limit);
    }

    [Theory]
    [InlineData("RtfImageCountLimitExceeded", 0, null, null)]
    [InlineData("RtfImagePayloadLimitExceeded", 1, 3, null)]
    [InlineData("RtfTotalImageLimitExceeded", 1, null, 3)]
    public void Image_Limits_Stop_Decoded_Hex_Payloads(string expectedCode, int? count, int? perImage, int? total) {
        const string rtf = @"{\rtf1{\pict\pngblip 01020304}}";
        var options = new RtfReadOptions {
            MaxImageCount = count,
            MaxImageBytesPerImage = perImage,
            MaxTotalImageBytes = total
        };

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Read(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
    }

    [Theory]
    [InlineData("RtfObjectCountLimitExceeded", 0, null, null)]
    [InlineData("RtfObjectPayloadLimitExceeded", 1, 3, null)]
    [InlineData("RtfTotalObjectLimitExceeded", 1, null, 3)]
    public void Object_Limits_Stop_Decoded_Hex_Payloads(string expectedCode, int? count, int? perObject, int? total) {
        const string rtf = @"{\rtf1{\object\objemb{\*\objdata 01020304}}}";
        var options = new RtfReadOptions {
            MaxObjectCount = count,
            MaxObjectBytesPerObject = perObject,
            MaxTotalObjectBytes = total
        };

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Read(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
    }

    [Fact]
    public void Semantic_Block_Limit_Stops_Binding() {
        const string rtf = @"{\rtf1 One\par Two\par Three\par}";
        var options = new RtfReadOptions { MaxSemanticBlockCount = 1 };

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.Read(rtf, options));

        Assert.Equal("RtfSemanticBlockLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadOptions.MaxSemanticBlockCount), exception.LimitSource);
    }

    [Fact]
    public void Untrusted_Profile_Blocks_Objects_And_File_References_With_Diagnostics() {
        const string rtf = @"{\rtf1{\filetbl{\file\fid0 C:\\private.txt;}}{\object\objemb{\*\objdata 0102}}Visible}";

        RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

        Assert.Empty(result.Document.FileReferences);
        Assert.Empty(result.Document.Blocks.OfType<RtfObject>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF105");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF106");
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Unknown_Ignorable_Destination_Is_Preserved_And_Diagnosed() {
        const string rtf = @"{\rtf1{\*\vendorprivate secret}Visible}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("Visible", result.Document.Paragraphs[0].ToPlainText());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("vendorprivate"));
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/private.txt")]
    public void Untrusted_Profile_Flattens_Unsafe_Hyperlink_Fields(string target) {
        string rtf = "{\\rtf1{\\field{\\*\\fldinst HYPERLINK \\\"" + target + "\\\"}{\\fldrslt Visible}}}";

        RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Empty(paragraph.Inlines.OfType<RtfField>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Untrusted_Profile_Preserves_Web_Hyperlink_Fields() {
        const string rtf = @"{\rtf1{\field{\*\fldinst HYPERLINK ""https://example.test/""}{\fldrslt Visible}}}";

        RtfReadResult result = RtfDocument.Read(rtf, RtfReadOptions.CreateUntrustedProfile());

        Assert.Single(Assert.Single(result.Document.Paragraphs).Inlines.OfType<RtfField>());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
    }

    [Fact]
    public void Synchronous_Read_Honors_Cancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            RtfDocument.Read(@"{\rtf1 Cancelled}", new RtfReadOptions(), cancellation.Token));
    }

    [Fact]
    public async Task Async_Read_Honors_Cancellation_During_Core_Pipeline() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            RtfDocument.ReadAsync(@"{\rtf1 Cancelled}", cancellationToken: cancellation.Token));
    }
}
