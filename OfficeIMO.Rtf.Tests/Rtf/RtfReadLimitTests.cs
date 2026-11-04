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
    public void Default_Profile_Is_Bounded_And_Blocks_External_Object_References() {
        var options = new RtfLoadOptions();

        Assert.Equal(128, options.Limits.MaxDepth);
        Assert.Equal(16L * 1024 * 1024, options.Limits.MaxInputBytes);
        Assert.Equal(16 * 1024 * 1024, options.Limits.MaxInputCharacters);
        Assert.Equal(1_000_000, options.Limits.MaxTokens);
        Assert.Equal(250_000, options.Limits.MaxGroups);
        Assert.Equal(8_000_000, options.Limits.MaxTextCharacters);
        Assert.Equal(4 * 1024 * 1024, options.Limits.MaxBinaryPayloadBytes);
        Assert.Equal(256, options.Limits.MaxImages);
        Assert.Equal(32, options.Limits.MaxObjects);
        Assert.Equal(100_000, options.Limits.MaxSemanticBlocks);
        Assert.False(options.ReadEmbeddedObjects);
        Assert.False(options.ReadFileReferences);
        Assert.Equal(RtfHyperlinkReadPolicy.WebAndMailOnly, options.HyperlinkPolicy);
    }

    [Fact]
    public void Compatibility_Profile_Explicitly_Restores_Legacy_Permissive_Behavior() {
        RtfLoadOptions options = RtfLoadOptions.CreateCompatibilityProfile();

        Assert.Equal(RtfReadLimits.CompatibilityMaxDepth, options.Limits.MaxDepth);
        Assert.Null(options.Limits.MaxInputBytes);
        Assert.Null(options.Limits.MaxInputCharacters);
        Assert.Null(options.Limits.MaxTokens);
        Assert.Null(options.Limits.MaxGroups);
        Assert.Null(options.Limits.MaxTextCharacters);
        Assert.Null(options.Limits.MaxSemanticBlocks);
        Assert.True(options.ReadEmbeddedObjects);
        Assert.True(options.ReadFileReferences);
        Assert.Equal(RtfHyperlinkReadPolicy.AllowAll, options.HyperlinkPolicy);
    }

    [Fact]
    public void Untrusted_Profile_Is_Bounded_And_Blocks_External_Object_References() {
        RtfLoadOptions options = RtfLoadOptions.UntrustedDefaults;

        Assert.Equal(128, options.Limits.MaxDepth);
        Assert.Equal(16L * 1024 * 1024, options.Limits.MaxInputBytes);
        Assert.Equal(1_000_000, options.Limits.MaxTokens);
        Assert.Equal(4 * 1024 * 1024, options.Limits.MaxBinaryPayloadBytes);
        Assert.Equal(256, options.Limits.MaxImages);
        Assert.Equal(32, options.Limits.MaxObjects);
        Assert.Equal(100_000, options.Limits.MaxSemanticBlocks);
        Assert.False(options.ReadEmbeddedObjects);
        Assert.False(options.ReadFileReferences);
        Assert.Equal(RtfHyperlinkReadPolicy.WebAndMailOnly, options.HyperlinkPolicy);
    }

    [Theory]
    [InlineData("RtfInputCharacterLimitExceeded", "MaxInputCharacters")]
    [InlineData("RtfTokenLimitExceeded", "MaxTokens")]
    [InlineData("RtfGroupLimitExceeded", "MaxGroups")]
    [InlineData("RtfTextCharacterLimitExceeded", "MaxTextCharacters")]
    public void Syntax_Limits_Throw_Stable_Exceptions(string expectedCode, string limitSource) {
        const string rtf = @"{\rtf1{nested}Visible text}";
        var options = new RtfLoadOptions();
        switch (limitSource) {
            case "MaxInputCharacters": options.Limits.MaxInputCharacters = 5; break;
            case "MaxTokens": options.Limits.MaxTokens = 2; break;
            case "MaxGroups": options.Limits.MaxGroups = 1; break;
            case "MaxTextCharacters": options.Limits.MaxTextCharacters = 3; break;
        }

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
        Assert.Equal(limitSource, exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void Byte_Limit_Is_Checked_Before_Seekable_Stream_Is_Read() {
        byte[] bytes = Encoding.ASCII.GetBytes(@"{\rtf1 Too large}");
        using var stream = new MemoryStream(bytes);
        var options = new RtfLoadOptions { Limits = { MaxInputBytes = 4 }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.LoadResult(stream, options));

        Assert.Equal("RtfInputByteLimitExceeded", exception.Code);
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public async Task Byte_Limit_Is_Enforced_By_Async_Stream_Load() {
        byte[] bytes = Encoding.ASCII.GetBytes(@"{\rtf1 Too large}");
        using var stream = new MemoryStream(bytes);
        var options = new RtfLoadOptions { Limits = { MaxInputBytes = 4 }};

        RtfReadLimitException exception = await Assert.ThrowsAsync<RtfReadLimitException>(() =>
            RtfDocument.LoadResultAsync(stream, options));

        Assert.Equal("RtfInputByteLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadLimits.MaxInputBytes), exception.LimitSource);
    }

    [Theory]
    [InlineData("RtfBinaryPayloadLimitExceeded", 3, null)]
    [InlineData("RtfTotalBinaryLimitExceeded", null, 3)]
    public void Binary_Limits_Are_Checked_Before_Payload_Allocation(string expectedCode, int? perPayload, int? total) {
        const string rtf = "{\\rtf1\\bin4 abcd}";
        var options = new RtfLoadOptions { Limits = {
            MaxBinaryPayloadBytes = perPayload,
            MaxTotalBinaryBytes = total }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

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
        var options = new RtfLoadOptions { Limits = {
            MaxImages = count,
            MaxImageBytes = perImage,
            MaxTotalImageBytes = total }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
    }

    [Theory]
    [InlineData("RtfObjectCountLimitExceeded", 0, null, null)]
    [InlineData("RtfObjectPayloadLimitExceeded", 1, 3, null)]
    [InlineData("RtfTotalObjectLimitExceeded", 1, null, 3)]
    public void Object_Limits_Stop_Decoded_Hex_Payloads(string expectedCode, int? count, int? perObject, int? total) {
        const string rtf = @"{\rtf1{\object\objemb{\*\objdata 01020304}}}";
        var options = new RtfLoadOptions { Limits = {
            MaxObjects = count,
            MaxObjectBytes = perObject,
            MaxTotalObjectBytes = total },
            ReadEmbeddedObjects = true};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal(expectedCode, exception.Code);
    }

    [Fact]
    public void Semantic_Block_Limit_Stops_Binding() {
        const string rtf = @"{\rtf1 One\par Two\par Three\par}";
        var options = new RtfLoadOptions { Limits = { MaxSemanticBlocks = 1 }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal("RtfSemanticBlockLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadLimits.MaxSemanticBlocks), exception.LimitSource);
    }

    [Fact]
    public void Semantic_Block_Limit_Counts_Paragraphs_Inside_Tables() {
        const string rtf = @"{\rtf1\trowd\cellx1000\intbl Cell\cell\row}";
        var options = new RtfLoadOptions { Limits = { MaxSemanticBlocks = 1 }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal("RtfSemanticBlockLimitExceeded", exception.Code);
        Assert.Equal(nameof(RtfReadLimits.MaxSemanticBlocks), exception.LimitSource);
    }

    [Theory]
    [InlineData(@"{\rtf1{\header Header\par}Body\par}")]
    [InlineData(@"{\rtf1 Body{\footnote Note\par}\par}")]
    [InlineData(@"{\rtf1{\shp{\*\shpinst{\shptxt Shape\par}}}}")]
    public void Semantic_Block_Limit_Counts_Paragraphs_In_Nested_Content(string rtf) {
        var options = new RtfLoadOptions { Limits = { MaxSemanticBlocks = 1 }};

        RtfReadLimitException exception = Assert.Throws<RtfReadLimitException>(() => RtfDocument.ParseResult(rtf, options));

        Assert.Equal("RtfSemanticBlockLimitExceeded", exception.Code);
    }

    [Fact]
    public void Untrusted_Profile_Blocks_Objects_And_File_References_With_Diagnostics() {
        const string rtf = @"{\rtf1{\filetbl{\file\fid0 C:\\private.txt;}}{\object\objemb{\*\objdata 0102}}Visible}";

        RtfReadResult result = RtfDocument.ParseResult(rtf, RtfLoadOptions.UntrustedDefaults);

        Assert.Empty(result.Document.FileReferences);
        Assert.Empty(result.Document.Blocks.OfType<RtfObject>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF105");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF106");
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Default_Profile_Blocks_Objects_File_References_And_Unsafe_Hyperlinks() {
        const string rtf = @"{\rtf1{\filetbl{\file\fid0 C:\\private.txt;}}{\object\objemb{\*\objdata 0102}}{\field{\*\fldinst HYPERLINK ""file:///C:/private.txt""}{\fldrslt Visible}}}";

        RtfReadResult result = RtfDocument.ParseResult(rtf);

        Assert.Empty(result.Document.FileReferences);
        Assert.Empty(result.Document.Blocks.OfType<RtfObject>());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Empty(paragraph.Inlines.OfType<RtfField>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF105");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF106");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Unknown_Ignorable_Destination_Is_Preserved_And_Diagnosed() {
        const string rtf = @"{\rtf1{\*\vendorprivate secret}Visible}";

        RtfReadResult result = RtfDocument.ParseResult(rtf);

        Assert.Equal("Visible", result.Document.Paragraphs[0].ToPlainText());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101" && diagnostic.Message.Contains("vendorprivate"));
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/private.txt")]
    public void Untrusted_Profile_Flattens_Unsafe_Hyperlink_Fields(string target) {
        string rtf = "{\\rtf1{\\field{\\*\\fldinst HYPERLINK \\\"" + target + "\\\"}{\\fldrslt Visible}}}";

        RtfReadResult result = RtfDocument.ParseResult(rtf, RtfLoadOptions.UntrustedDefaults);

        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Empty(paragraph.Inlines.OfType<RtfField>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
        Assert.Equal(rtf, result.ToRtfLossless());
    }

    [Fact]
    public void Untrusted_Profile_Preserves_Web_Hyperlink_Fields() {
        const string rtf = @"{\rtf1{\field{\*\fldinst HYPERLINK ""https://example.test/""}{\fldrslt Visible}}}";

        RtfReadResult result = RtfDocument.ParseResult(rtf, RtfLoadOptions.UntrustedDefaults);

        Assert.Single(Assert.Single(result.Document.Paragraphs).Inlines.OfType<RtfField>());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
    }

    [Theory]
    [InlineData("../payload.exe")]
    [InlineData("payload.exe")]
    [InlineData("\\\\server\\share\\payload.exe")]
    public void Default_Profile_Flattens_Relative_And_FileLike_Hyperlink_Targets(string target) {
        RtfDocument authored = RtfDocument.Create();
        RtfField authoredField = authored.AddParagraph().AddField(
            new RtfHyperlinkFieldInfo { Target = new Uri(target, UriKind.RelativeOrAbsolute) }.ToInstruction());
        authoredField.Result.AddText("Visible");
        string rtf = authored.ToRtf();

        RtfReadResult result = RtfDocument.ParseResult(rtf);

        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible", paragraph.ToPlainText());
        Assert.Empty(paragraph.Inlines.OfType<RtfField>());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
    }

    [Fact]
    public void Default_Profile_Preserves_SubaddressOnly_Hyperlinks() {
        RtfDocument authored = RtfDocument.Create();
        RtfField authoredField = authored.AddParagraph().AddField(
            new RtfHyperlinkFieldInfo { SubAddress = "bookmark" }.ToInstruction());
        authoredField.Result.AddText("Jump");

        RtfReadResult result = RtfDocument.ParseResult(authored.ToRtf());

        RtfField field = Assert.Single(Assert.Single(result.Document.Paragraphs).Inlines.OfType<RtfField>());
        Assert.Null(field.HyperlinkField!.Target);
        Assert.Equal("bookmark", field.HyperlinkField.SubAddress);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF107");
    }

    [Fact]
    public void Synchronous_Read_Honors_Cancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            RtfDocument.ParseResult(@"{\rtf1 Cancelled}", new RtfLoadOptions(), cancellation.Token));
    }

}
