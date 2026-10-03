using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests.Emlx;

public sealed class EmlxMetadataPreservationTests {
    // Produced by Python plistlib, then independently accepted by macOS plutil -lint.
    private const string BinaryFixture = "YnBsaXN0MDDaAQIDBAUGBwgJCgsMDQ4TFBUWFxhVZmxhZ3NWdmVuZG9yVlZlbmRvclZuZXN0ZWRXdW5pY29kZVVieXRlc1R0aW1lWG5lZ2F0aXZlVm51bWJlclRsaXN0EwAAAQAAAAABU29uZVN0d2/SDxARElNLZXlTa2V5UWFRYmoAWgBhAXwA8wFCAQcAIGXlZyyKnkMAAQIzQcg36iAAAAAT//////////4jP/gAAAAAAACjGRobCRADUXgIHSMqMThARktUW2BpbXF2en6AgpebpK22uru9AAAAAAAAAQEAAAAAAAAAHAAAAAAAAAAAAAAAAAAAAL8=";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BinaryAndXmlMetadataKeepExactKeysAndUnknownFlagsThroughEditedRewrite(bool binary) {
        byte[] metadata = binary ? Convert.FromBase64String(BinaryFixture) : Encoding.UTF8.GetBytes(
            "<plist><dict><key>flags</key><integer>1099511627777</integer>" +
            "<key>vendor</key><string>one</string><key>Vendor</key><string>two</string>" +
            "<key>nested</key><dict><key>Key</key><string>a</string><key>key</key><string>b</string></dict></dict></plist>");
        EmailDocument original = Read(metadata, out EmailStoreReadResult read);
        Assert.Empty(read.Diagnostics);
        Assert.True(original.MessageMetadata.IsRead);
        original.MessageMetadata.IsRead = false;
        original.Subject = "Edited subject";
        using var portable = new MemoryStream();
        EmailWriteResult portableResult = new EmailDocumentWriter().Write(original, portable);
        Assert.Equal(EmailConversionLossDisposition.Accepted, portableResult.LossDisposition);
        Assert.Contains(portableResult.Diagnostics, item => item.Code == "EMAIL_EMLX_METADATA_NOT_REPRESENTED");
        Assert.Throws<InvalidDataException>(() => portableResult.RequireNoLoss());
        using var output = new MemoryStream();
        EmailWriteResult result = new EmailStoreEmlxWriter().Write(original, output);
        result.RequireNoLoss();
        EmailDocument rewritten = new EmailStoreReader().Read(output, "rewritten.emlx").Store.Folders.Single().Items.Single().Document;

        IReadOnlyDictionary<string, object?> values = Metadata(rewritten);
        Assert.Equal(1L << 40, values["flags"]);
        Assert.Equal("one", values["vendor"]);
        Assert.Equal("two", values["Vendor"]);
        IReadOnlyDictionary<string, object?> nested = Assert.IsAssignableFrom<IReadOnlyDictionary<string, object?>>(values["nested"]);
        Assert.Equal("a", nested["Key"]);
        Assert.Equal("b", nested["key"]);
        Assert.Equal("Edited subject", rewritten.Subject);
        Assert.False(rewritten.MessageMetadata.IsRead);
        if (binary) {
            Assert.Equal("Zażółć 日本語", values["unicode"]);
            Assert.Equal(new byte[] { 0, 1, 2 }, Assert.IsType<byte[]>(values["bytes"]));
            Assert.Equal(-2L, values["negative"]);
            Assert.Equal(1.5, values["number"]);
            Assert.Equal(new DateTimeOffset(2026, 10, 2, 12, 0, 0, TimeSpan.Zero), values["time"]);
            Assert.Equal(new object?[] { true, 3L, "x" }, Assert.IsType<object?[]>(values["list"]));
        }
    }

    [Fact]
    public void ExactMetadataEditsAndRemovalOverrideReadOnlyFlatAliases() {
        EmailDocument document = Read(Encoding.UTF8.GetBytes(
            "<plist><dict><key>vendor</key><string>old</string><key>remove</key><string>old</string><key>flags</key><integer>1099511627777</integer></dict></plist>"), out _);
        var exact = new Dictionary<string, object?>(Metadata(document), StringComparer.Ordinal) {
            ["vendor"] = "new"
        };
        exact.Remove("remove");
        exact.Remove("flags");
        document.Properties["Emlx:Metadata"] = exact;
        using var output = new MemoryStream();
        new EmailStoreEmlxWriter().Write(document, output).RequireNoLoss();
        EmailDocument rewritten = new EmailStoreReader().Read(output, "edited.emlx")
            .Store.Folders.Single().Items.Single().Document;
        Assert.Equal("new", Metadata(rewritten)["vendor"]);
        Assert.False(Metadata(rewritten).ContainsKey("remove"));
        Assert.Equal(1L, Metadata(rewritten)["flags"]);
    }

    [Theory]
    [InlineData(EmailConversionLossPolicy.Block)]
    [InlineData(EmailConversionLossPolicy.Warn)]
    [InlineData(EmailConversionLossPolicy.Allow)]
    public void OpaqueMetadataRequiresExplicitLossPolicyAndRetainsItsExactBytes(EmailConversionLossPolicy policy) {
        byte[] raw = Encoding.UTF8.GetBytes("bplist00invalid-trailer");
        EmailDocument document = Read(raw, out _);
        var writer = new EmailStoreEmlxWriter(new EmailStoreEmlxWriterOptions(
            messageOptions: new EmailWriterOptions(conversionLossPolicy: policy)));
        using var output = new MemoryStream();
        output.WriteByte(123);
        EmailWriteResult result = writer.Write(document, output);

        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_EMLX_METADATA_OPAQUE");
        Assert.Throws<InvalidDataException>(() => result.RequireNoLoss());
        if (policy == EmailConversionLossPolicy.Block) {
            Assert.Equal(EmailConversionLossDisposition.Blocked, result.LossDisposition);
            Assert.Equal(new byte[] { 123 }, output.ToArray());
        } else {
            Assert.Equal(EmailConversionLossDisposition.Accepted, result.LossDisposition);
            EmailDocument loaded = new EmailStoreReader().Read(output, "opaque.emlx").Store.Folders.Single().Items.Single().Document;
            Assert.EndsWith(Encoding.UTF8.GetString(raw), Encoding.UTF8.GetString(Assert.IsType<byte[]>(loaded.Properties["Emlx:RawMetadata"])));
        }
    }

    [Fact]
    public void BinaryMetadataObjectAndPropertyLimitsStopBeforeProjection() {
        byte[] metadata = Convert.FromBase64String(BinaryFixture);
        using var input = Envelope(metadata);
        EmailStoreLimitExceededException error = Assert.Throws<EmailStoreLimitExceededException>(() =>
            new EmailStoreReader(new EmailStoreReaderOptions(maxPropertiesPerItem: 2)).Read(input, "bounded.emlx"));
        Assert.Equal(nameof(EmailStoreReaderOptions.MaxPropertiesPerItem), error.LimitName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void InvalidBinaryReferencesOffsetsAndCyclesDoNotDiscardTheMessage(int mutation) {
        byte[] metadata = Convert.FromBase64String(BinaryFixture);
        int trailer = metadata.Length - 32;
        if (mutation == 0) metadata[trailer + 16] = 255; // root outside object table
        else if (mutation == 1) metadata[trailer + 24] = 255; // offset table outside source
        else metadata[19] = 0; // root dictionary's first value points back to the root
        EmailDocument document = Read(metadata, out EmailStoreReadResult result);
        Assert.Equal("Original", document.Subject);
        Assert.Contains(result.Diagnostics, item => item.Code == "EMAIL_STORE_EMLX_METADATA_INVALID");
        Assert.Equal(true, document.Properties["Emlx:MetadataOpaque"]);
    }

    private static IReadOnlyDictionary<string, object?> Metadata(EmailDocument document) =>
        Assert.IsAssignableFrom<IReadOnlyDictionary<string, object?>>(document.Properties["Emlx:Metadata"]);

    private static EmailDocument Read(byte[] metadata, out EmailStoreReadResult result) {
        using var input = Envelope(metadata);
        result = new EmailStoreReader().Read(input, "fixture.emlx");
        return Assert.Single(Assert.Single(result.Store.Folders).Items).Document;
    }

    private static MemoryStream Envelope(byte[] metadata) {
        byte[] message = Encoding.UTF8.GetBytes("Subject: Original\r\n\r\nBody\r\n");
        var stream = new MemoryStream();
        byte[] prefix = Encoding.ASCII.GetBytes(message.Length + "\n");
        stream.Write(prefix, 0, prefix.Length);
        stream.Write(message, 0, message.Length);
        stream.Write(metadata, 0, metadata.Length);
        stream.Position = 0;
        return stream;
    }
}
