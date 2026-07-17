using OfficeIMO.Drawing.Internal;
using OpenMcdf;
using System.Threading;
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
    public void SerializedLengthPreflightMatchesCompoundOutput() {
        var streams = new[] {
            new OfficeCompoundStream("Small", new byte[17]),
            new OfficeCompoundStream("Storage/Regular", new byte[5000]),
            new OfficeCompoundStream("Storage/Empty", Array.Empty<byte>())
        };

        long expectedLength = OfficeCompoundFileWriter.GetSerializedLength(streams);
        byte[] compound = OfficeCompoundFileWriter.Write(streams);

        Assert.Equal(compound.LongLength, expectedLength);
    }

    [Fact]
    public void VersionThreeWriterRejectsAStreamLargerThanItsDirectoryCanRepresent() {
        var streams = new[] {
            new OfficeCompoundStream("TooLarge", (long)uint.MaxValue + 1,
                () => throw new InvalidOperationException("Preflight must reject before opening content."))
        };

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            OfficeCompoundFileWriter.GetSerializedLength(streams));

        Assert.Contains("4 GiB", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SelectiveReaderExternalizesARegularStreamWithoutMaterializingIt() {
        byte[] large = Enumerable.Range(0, 1024 * 1024 + 19).Select(index => (byte)(index % 251)).ToArray();
        byte[] compoundBytes = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("Metadata", Encoding.UTF8.GetBytes("small")),
            new OfficeCompoundStream("Storage/Large", large)
        });
        string externalPath = Path.GetTempFileName();
        try {
            using var input = new MemoryStream(new byte[7].Concat(compoundBytes).ToArray());
            input.Position = 7;
            long originalPosition = input.Position;
            bool success = OfficeCompoundFileReader.TryReadSelective(input,
                new OfficeCompoundReadOptions(maxStreamBytes: 2 * 1024 * 1024),
                (path, _) => path == "Storage/Large",
                (_, _) => new FileStream(externalPath, FileMode.Create, FileAccess.Write, FileShare.Read),
                out OfficeCompoundFile? compound, out string? error);

            Assert.True(success, error);
            Assert.Equal(originalPosition, input.Position);
            Assert.Equal("small", Encoding.UTF8.GetString(compound!.Streams["Metadata"]));
            Assert.Empty(compound.Streams["Storage/Large"]);
            Assert.Equal(large, File.ReadAllBytes(externalPath));
        } finally {
            File.Delete(externalPath);
        }
    }

    [Fact]
    public void SelectiveReaderObservesCancellationFromStreamCallback() {
        byte[] compoundBytes = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("Empty", Array.Empty<byte>())
        });
        using var input = new MemoryStream(compoundBytes);
        using var cancellation = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeCompoundFileReader.TryReadSelective(input,
                new OfficeCompoundReadOptions(),
                (_, _) => {
                    cancellation.Cancel();
                    return false;
                },
                (_, _) => throw new InvalidOperationException(
                    "No stream should be externalized."),
                cancellation.Token, out _, out _));

        Assert.Equal(0, input.Position);
    }

    [Fact]
    public void RewritePreservesRetainedMetadataAndRemovesStorageSubtree() {
        var rootClassId = new Guid(
            "64818D10-4F9B-11CF-86EA-00AA00B929E8");
        var storageClassId = new Guid(
            "00020906-0000-0000-C000-000000000046");
        var entries = new[] {
            new OfficeCompoundFileEntry("Keep", "Keep", 1, 0,
                classId: storageClassId, stateBits: 9,
                creationTime: 11, modifiedTime: 12),
            new OfficeCompoundFileEntry("Value", "Keep/Value", 2, 3),
            new OfficeCompoundFileEntry("Remove", "Remove", 1, 0),
            new OfficeCompoundFileEntry("Value", "Remove/Value", 2, 3)
        };
        var source = new OfficeCompoundFile(
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase) {
                ["Keep/Value"] = new byte[] { 1, 2, 3 },
                ["Remove/Value"] = new byte[] { 4, 5, 6 }
            },
            entries,
            new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5,
                0, classId: rootClassId));

        byte[] rewritten = OfficeCompoundFileWriter.Rewrite(source,
            new Dictionary<string, byte[]> {
                ["Keep/Value"] = new byte[] { 7, 8, 9 }
            },
            new[] { "Remove" });

        Assert.True(OfficeCompoundFileReader.TryRead(rewritten,
            out OfficeCompoundFile? result, out string? error), error);
        Assert.Equal(new byte[] { 7, 8, 9 },
            result!.Streams["Keep/Value"]);
        Assert.DoesNotContain(result.Entries, entry =>
            entry.Path.StartsWith("Remove", StringComparison.Ordinal));
        OfficeCompoundFileEntry retained = Assert.Single(result.Entries,
            entry => entry.Path == "Keep" && !entry.IsFallback);
        Assert.Equal(storageClassId, retained.ClassId);
        Assert.Equal(9U, retained.StateBits);
        Assert.Equal(11UL, retained.CreationTime);
        Assert.Equal(12UL, retained.ModifiedTime);
        Assert.Equal(rootClassId, result.RootEntry.ClassId);
    }

    [Fact]
    public void WriterEmitsDifatForLargeCompoundFiles() {
        byte[] content = new byte[8 * 1024 * 1024];
        for (int i = 0; i < content.Length; i += 4096) content[i] = (byte)(i / 4096 % 251);

        var streams = new[] { new OfficeCompoundStream("Large", content) };
        long expectedLength = OfficeCompoundFileWriter.GetSerializedLength(streams);
        byte[] compound = OfficeCompoundFileWriter.Write(streams);
        uint fatSectorCount = ReadUInt32(compound, 44);
        uint firstDifatSector = ReadUInt32(compound, 68);
        bool success = OfficeCompoundFileReader.TryRead(compound, out OfficeCompoundFile? file, out string? error);

        Assert.True(fatSectorCount > 109);
        Assert.NotEqual(0xfffffffeU, firstDifatSector);
        Assert.Equal(expectedLength, compound.LongLength);
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
    public void EmailPolicyBoundsDeclaredRootMiniStreamBeforeMaterializingIt() {
        byte[] compound = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("Property", new byte[] { 1 })
        });
        int directoryOffset = checked(((int)ReadUInt32(compound, 48) + 1) * 512);
        WriteUInt64(compound, directoryOffset + 120, 512);
        var readerOptions = new EmailReaderOptions(
            maxInputBytes: 8192,
            maxCompoundDirectoryEntries: 4,
            maxDecodedPropertyBytes: 1,
            maxTotalAttachmentBytes: 1);

        bool success = OfficeCompoundFileReader.TryRead(compound,
            EmailCompoundReadPolicy.Create(readerOptions), out _, out string? error);

        Assert.False(success);
        Assert.Contains("mini stream exceeds configured bounds", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SelectiveReaderBoundsDeclaredRootMiniStreamBeforeBuildingItsSectorIndex() {
        byte[] compound = OfficeCompoundFileWriter.Write(new[] {
            new OfficeCompoundStream("Property", new byte[] { 1 })
        });
        int directoryOffset = checked(((int)ReadUInt32(compound, 48) + 1) * 512);
        WriteUInt64(compound, directoryOffset + 120, 512);
        var options = new OfficeCompoundReadOptions(
            maxTotalStreamBytes: 1);

        using var source = new MemoryStream(compound);
        bool success = OfficeCompoundFileReader.TryReadSelective(source, options,
            (_, _) => false,
            (_, _) => throw new InvalidOperationException("No stream should be externalized."),
            out _, out string? error);

        Assert.False(success);
        Assert.Contains("mini stream exceeds configured or physical bounds", error,
            StringComparison.OrdinalIgnoreCase);
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

    private static void WriteUInt64(byte[] bytes, int offset, ulong value) {
        for (int index = 0; index < 8; index++) {
            bytes[offset + index] = (byte)(value >> (index * 8));
        }
    }
}
