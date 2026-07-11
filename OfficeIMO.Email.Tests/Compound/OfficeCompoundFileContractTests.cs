using OfficeIMO.Shared;
using OpenMcdf;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class OfficeCompoundFileContractTests {
    [Fact]
    public void HierarchicalWriterRoundTripsMiniRegularAndEmptyStreams() {
        byte[] regular = Enumerable.Range(0, 5000).Select(index => (byte)(index % 251)).ToArray();
        var streams = new[] {
            new OfficeCompoundStream("Top", new byte[] { 1, 2, 3 }),
            new OfficeCompoundStream("Storage/Nested/Small", Encoding.UTF8.GetBytes("payload")),
            new OfficeCompoundStream("Storage/Regular", regular),
            new OfficeCompoundStream("Storage/Empty", Array.Empty<byte>())
        };

        byte[] first = OfficeCompoundFileWriter.Write(streams);
        byte[] second = OfficeCompoundFileWriter.Write(streams);
        bool success = OfficeCompoundFileReader.TryRead(first, out OfficeCompoundFile? file, out string? error);

        Assert.True(success, error);
        Assert.Equal(first, second);
        Assert.Equal(new byte[] { 1, 2, 3 }, file!.Streams["Top"]);
        Assert.Equal("payload", Encoding.UTF8.GetString(file.Streams["Storage/Nested/Small"]));
        Assert.Equal(regular, file.Streams["Storage/Regular"]);
        Assert.Empty(file.Streams["Storage/Empty"]);
        Assert.Contains(file.Entries, entry => entry.IsStorage && entry.Path == "Storage/Nested");

        using MemoryStream source = new MemoryStream(first);
        using RootStorage oracle = RootStorage.Open(source, StorageModeFlags.LeaveOpen);
        Storage storage = oracle.OpenStorage("Storage");
        Storage nested = storage.OpenStorage("Nested");
        using CfbStream small = nested.OpenStream("Small");
        using StreamReader text = new StreamReader(small, Encoding.UTF8);
        Assert.Equal("payload", text.ReadToEnd());
    }

    [Fact]
    public void WriterPersistsAnExplicitRootStorageClassId() {
        var classId = new Guid("00020D0B-0000-0000-C000-000000000046");

        byte[] compound = OfficeCompoundFileWriter.Write(
            new[] { new OfficeCompoundStream("Payload", new byte[] { 1, 2, 3 }) },
            classId);

        using MemoryStream source = new MemoryStream(compound);
        using RootStorage oracle = RootStorage.Open(source, StorageModeFlags.LeaveOpen);
        Assert.Equal(classId, oracle.EntryInfo.CLSID);
    }

    [Fact]
    public void WriterEmitsDifatForLargeCompoundFiles() {
        byte[] content = new byte[8 * 1024 * 1024];
        for (int i = 0; i < content.Length; i += 4096) content[i] = (byte)(i / 4096 % 251);

        byte[] compound = OfficeCompoundFileWriter.Write(new[] { new OfficeCompoundStream("Large", content) });
        uint fatSectorCount = ReadUInt32(compound, 44);
        uint firstDifatSector = ReadUInt32(compound, 68);
        bool success = OfficeCompoundFileReader.TryRead(compound, out OfficeCompoundFile? file, out string? error);

        Assert.True(fatSectorCount > 109);
        Assert.NotEqual(0xfffffffeU, firstDifatSector);
        Assert.True(success, error);
        Assert.Equal(content, file!.Streams["Large"]);

        using MemoryStream source = new MemoryStream(compound);
        using RootStorage oracle = RootStorage.Open(source, StorageModeFlags.LeaveOpen);
        using CfbStream large = oracle.OpenStream("Large");
        Assert.Equal(content.Length, large.Length);
    }

    [Fact]
    public void ReaderRejectsAggregateExpansionBeyondConfiguredLimit() {
        byte[] compound = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("One", new byte[] { 1, 2, 3 }),
            new OfficeCompoundStream("Two", new byte[] { 4, 5, 6 })
        });
        var options = new OfficeCompoundReadOptions(maxTotalStreamBytes: 5);

        bool success = OfficeCompoundFileReader.TryRead(compound, options, out _, out string? error);

        Assert.False(success);
        Assert.Contains("exceed", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AttachmentPolicyRejectsDeclaredStreamTotalsBeforeMaterializingPayloads() {
        byte[] compound = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("One", new byte[400]),
            new OfficeCompoundStream("Two", new byte[400])
        });
        var readerOptions = new EmailReaderOptions(maxAttachmentBytes: 600);
        OfficeCompoundReadOptions compoundOptions =
            EmailCompoundReadPolicy.CreateForAttachment(readerOptions, existingTotalAttachmentBytes: 0);

        OfficeCompoundStreamLimitExceededException exception =
            Assert.Throws<OfficeCompoundStreamLimitExceededException>(() =>
                OfficeCompoundFileReader.TryRead(compound, compoundOptions, out _, out _));

        Assert.Equal(nameof(EmailReaderOptions.MaxAttachmentBytes), exception.LimitName);
        Assert.Equal(800, exception.ActualValue);
    }

    [Fact]
    public void ReaderRejectsDirectoryChainsBeforeBufferingBeyondTheConfiguredLimit() {
        byte[] compound = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("One", new byte[] { 1 })
        });
        var options = new OfficeCompoundReadOptions(maxDirectoryEntries: 1);

        bool success = OfficeCompoundFileReader.TryRead(compound, options, out _, out string? error);

        Assert.False(success);
        Assert.Contains("directory entry count exceeds 1", error, StringComparison.OrdinalIgnoreCase);
    }

    private static uint ReadUInt32(byte[] bytes, int offset) {
        return (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
    }
}
