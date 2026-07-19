using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

namespace OfficeIMO.Email.AddressBook.Tests;

public sealed class ReaderEmailAddressBookTests {
    [Fact]
    public void ItemReaderProjectsOneBoundedTypedChunkPerEntry() {
        using (var stream = new MemoryStream(new OabV4Fixture().Build(), writable: false)) {
            ReaderEmailAddressBookEntryResult[] results = EmailAddressBookEntryReader.Read(
                stream,
                "synthetic.oab",
                new ReaderOptions { ComputeHashes = true },
                new ReaderEmailAddressBookOptions { MaxEntries = 2 }).ToArray();

            Assert.Equal(2, results.Length);
            Assert.All(results, result => Assert.True(result.Succeeded));
            ReaderChunk chunk = Assert.Single(results[0].Chunks);
            Assert.Equal(ReaderInputKind.Email, chunk.Kind);
            Assert.Equal("oab:0000:0000000000", chunk.Id);
            Assert.Contains("Ada Lovelace", chunk.Text, StringComparison.Ordinal);
            Assert.Contains("ada@example.test", chunk.Markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("Binary", chunk.Text, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
        }
    }

    [Fact]
    public void MembershipValuesAreExplicitlyOptIn() {
        byte[] oab = new OabV4Fixture().Build();
        var query = new OfflineAddressBookSearchQuery(
            new[] { "All Example" },
            fields: OfflineAddressBookSearchFields.Names,
            objectType: OfflineAddressBookObjectType.DistributionList);

        using (var stream = new MemoryStream(oab, writable: false)) {
            ReaderChunk chunk = Assert.Single(Assert.Single(EmailAddressBookEntryReader.Read(
                stream,
                "synthetic.oab",
                addressBookOptions: new ReaderEmailAddressBookOptions { Query = query })).Chunks);
            Assert.Contains("Distribution-list member count: 2", chunk.Text, StringComparison.Ordinal);
            Assert.DoesNotContain("cn=ada", chunk.Text, StringComparison.Ordinal);
        }

        using (var stream = new MemoryStream(oab, writable: false)) {
            ReaderChunk chunk = Assert.Single(Assert.Single(EmailAddressBookEntryReader.Read(
                stream,
                "synthetic.oab",
                addressBookOptions: new ReaderEmailAddressBookOptions {
                    Query = query,
                    IncludeMembershipValues = true
                })).Chunks);
            Assert.Contains("cn=ada", chunk.Text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RegisteredHandlerProvidesChunksAndNativeDocumentEnvelope() {
        byte[] oab = new OabV4Fixture().Build();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddEmailAddressBookHandler(new ReaderEmailAddressBookOptions { MaxEntries = 1 })
            .Build();
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item =>
            item.Id == OfficeDocumentReaderBuilderEmailAddressBookExtensions.HandlerId);
        Assert.Equal(new[] { ".oab" }, capability.Extensions);
        Assert.True(capability.SupportsDocumentPath);
        Assert.True(capability.SupportsDocumentStream);

        using (var stream = new MemoryStream(oab, writable: false)) {
            OfficeDocumentReadResult result = reader.ReadDocument(
                stream, "synthetic.oab", new ReaderOptions { ComputeHashes = true });

            Assert.Equal(ReaderInputKind.Email, result.Kind);
            Assert.Single(result.Chunks);
            Assert.Contains(OfficeDocumentReaderBuilderEmailAddressBookExtensions.HandlerId,
                result.CapabilitiesUsed);
            Assert.Contains(result.Metadata, item =>
                item.Name == "ProjectedEntryCount" && item.Value == "1");
            Assert.Null(result.Source.SourceHash);
            Assert.False(string.IsNullOrWhiteSpace(result.Chunks[0].ChunkHash));
            Assert.Equal(0, stream.Position);
        }
    }
}
