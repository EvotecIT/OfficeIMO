using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives.Tests;

public sealed class CabinetCoreTests {
    [Fact]
    public void MsZipUsesThePreceding32KiBHistoryAcrossFrames() {
        byte[] first = Enumerable.Range(0, 32768).Select(index => (byte)(index % 251)).ToArray();
        byte[] second = first.Take(258).ToArray();
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(first, final: true), DictionaryMatchFrame() },
            new[] { first.Length, second.Length });

        OfficeCabinetEntry entry = Assert.Single(Read(cabinet));

        Assert.Equal("templates.admx", entry.Name);
        Assert.Equal(first.Concat(second), entry.Data);
    }

    [Fact]
    public void RejectsMissingFinalDeflateBlockEvenWhenAllAdvertisedBytesExist() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[] { 1, 2, 3, 4 }, final: false) }, new[] { 4 });

        OfficeCabinetException exception = Assert.Throws<OfficeCabinetException>(() => Read(cabinet));

        Assert.Equal("CAB_MSZIP_CORRUPT", exception.Code);
    }

    [Fact]
    public void RejectsTrailingDeflateData() {
        byte[] frame = StoredFrame(new byte[] { 1, 2, 3, 4 }, final: true).Concat(new byte[] { 0 }).ToArray();
        byte[] cabinet = BuildCabinet(new[] { frame }, new[] { 4 });

        Assert.Equal("CAB_MSZIP_CORRUPT", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Fact]
    public void RejectsChecksumMismatchBeforeDecoding() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[] { 1, 2, 3, 4 }, final: true) }, new[] { 4 });
        cabinet[cabinet.Length - 1] ^= 1;

        Assert.Equal("CAB_CHECKSUM", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Fact]
    public void FilteringDoesNotBypassDecodedFolderExpansionLimit() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[32], final: true) }, new[] { 32 }, declaredFileLength: 1);

        OfficeCabinetException exception = Assert.Throws<OfficeCabinetException>(() =>
            OfficeCabinetArchiveReader.Read(cabinet, 16, 16, 10, entryFilter: _ => false));

        Assert.Equal("CAB_EXPANDED_LIMIT", exception.Code);
    }

    [Fact]
    public void ReportsActualExpandedFolderBytesWhenNoMembersAreSelected() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[32], final: true) }, new[] { 32 }, declaredFileLength: 1);
        using var stream = new MemoryStream(cabinet, writable: false);

        IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(stream, 1024,
            1024, 1024, 10, out long expandedBytes, entryFilter: _ => false);

        Assert.Empty(entries);
        Assert.Equal(32, expandedBytes);
    }

    [Fact]
    public void MetadataValidationPrecedesSelectionAndRestoresPositionWhenItRejectsAMember() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[4], final: true) }, new[] { 4 });
        using var stream = new MemoryStream(cabinet, writable: false);
        stream.Position = 9;
        bool selected = false;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => OfficeCabinetArchiveReader.Read(
            stream, 1024, 1024, 1024, 10, out _, CancellationToken.None,
            _ => { selected = true; return false; }, null,
            (name, length) => {
                Assert.Equal("templates.admx", name); Assert.Equal(4, length);
                throw new InvalidDataException("metadata mismatch");
            }));

        Assert.Equal("metadata mismatch", exception.Message);
        Assert.False(selected);
        Assert.Equal(9, stream.Position);
    }

    [Fact]
    public void RejectsDeclaredEntryLimitBeforeDecoding() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[32], final: true) }, new[] { 32 });

        Assert.Equal("CAB_ENTRY_LIMIT", Assert.Throws<OfficeCabinetException>(() =>
            OfficeCabinetArchiveReader.Read(cabinet, 1024, 16, 10)).Code);
    }

    [Fact]
    public void RejectsInputLimitBeforeReadingAndPreservesCallerPosition() {
        using var stream = new RejectReadStream(new byte[128]);
        stream.Position = 9;

        Assert.Equal("CAB_INPUT_LIMIT", Assert.Throws<OfficeCabinetException>(() =>
            OfficeCabinetArchiveReader.Read(stream, 64, 1024, 1024, 10)).Code);
        Assert.Equal(9, stream.Position);
    }

    [Fact]
    public void CancellationDuringInputCopyRestoresCallerPosition() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[100], final: true) }, new[] { 100 });
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterReadStream(cabinet, cancellation);
        stream.Position = 9;

        Assert.Throws<OperationCanceledException>(() =>
            OfficeCabinetArchiveReader.Read(stream, 1024, 1024, 1024, 10, cancellation.Token));
        Assert.Equal(9, stream.Position);
    }

    [Fact]
    public void RejectedFilteredMemberStillHasItsRangeValidated() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[4], final: true) }, new[] { 4 }, declaredFileLength: 5);

        Assert.Equal("CAB_ENTRY_RANGE", Assert.Throws<OfficeCabinetException>(() =>
            OfficeCabinetArchiveReader.Read(cabinet, 1024, 1024, 10, entryFilter: _ => false)).Code);
    }

    [Fact]
    public void RejectsContinuedCabinetBeforeReadingAdditionalNames() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[4], final: true) }, new[] { 4 });
        UInt16(cabinet, 30, 2);

        Assert.Equal("CAB_MULTIPART", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    [Fact]
    public void DoesNotReadCabinetRecordsFromAnAppendedSignature() {
        byte[] cabinet = BuildCabinet(new[] { StoredFrame(new byte[4], final: true) }, new[] { 4 });
        UInt32(cabinet, 8, (uint)(cabinet.Length - 1));

        Assert.Equal("CAB_TRUNCATED", Assert.Throws<OfficeCabinetException>(() => Read(cabinet)).Code);
    }

    private static IReadOnlyList<OfficeCabinetEntry> Read(byte[] cabinet) =>
        OfficeCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10);

    private static byte[] BuildCabinet(IReadOnlyList<byte[]> frames, IReadOnlyList<int> sizes, uint? declaredFileLength = null) {
        byte[] name = Encoding.ASCII.GetBytes("templates.admx\0");
        const int filesOffset = 44;
        int dataOffset = filesOffset + 16 + name.Length;
        var cabinet = new byte[dataOffset + frames.Sum(frame => 8 + frame.Length)];
        cabinet[0] = (byte)'M'; cabinet[1] = (byte)'S'; cabinet[2] = (byte)'C'; cabinet[3] = (byte)'F';
        UInt32(cabinet, 8, (uint)cabinet.Length);
        UInt32(cabinet, 16, filesOffset);
        cabinet[24] = 3; cabinet[25] = 1;
        UInt16(cabinet, 26, 1);
        UInt16(cabinet, 28, 1);
        UInt32(cabinet, 36, (uint)dataOffset);
        UInt16(cabinet, 40, (ushort)frames.Count);
        UInt16(cabinet, 42, 1);
        UInt32(cabinet, filesOffset, declaredFileLength ?? (uint)sizes.Sum());
        Buffer.BlockCopy(name, 0, cabinet, filesOffset + 16, name.Length);
        int offset = dataOffset;
        for (int index = 0; index < frames.Count; index++) {
            byte[] frame = frames[index];
            UInt16(cabinet, offset + 4, (ushort)frame.Length);
            UInt16(cabinet, offset + 6, (ushort)sizes[index]);
            Buffer.BlockCopy(frame, 0, cabinet, offset + 8, frame.Length);
            UInt32(cabinet, offset, Checksum(cabinet, offset + 4, frame.Length + 4));
            offset += frame.Length + 8;
        }
        return cabinet;
    }

    private static byte[] StoredFrame(byte[] payload, bool final) {
        var frame = new byte[payload.Length + 7];
        frame[0] = 0x43; frame[1] = 0x4B; frame[2] = final ? (byte)1 : (byte)0;
        UInt16(frame, 3, (ushort)payload.Length);
        UInt16(frame, 5, unchecked((ushort)~payload.Length));
        Buffer.BlockCopy(payload, 0, frame, 7, payload.Length);
        return frame;
    }

    private static byte[] DictionaryMatchFrame() {
        // One fixed-Huffman match: length 258, distance 32768, followed by EOB.
        // It cannot decode without the first CFDATA frame's dictionary.
        var bits = new List<int>();
        AppendBits(bits, 1, 1);
        AppendBits(bits, 1, 2);
        AppendCode(bits, 197, 8); // fixed literal/length symbol 285
        AppendCode(bits, 29, 5); // distance 29, base 24577
        AppendBits(bits, 8191, 13);
        AppendCode(bits, 0, 7); // end of block, symbol 256
        var result = new byte[2 + (bits.Count + 7) / 8];
        result[0] = 0x43; result[1] = 0x4B;
        for (int index = 0; index < bits.Count; index++) result[2 + index / 8] |= (byte)(bits[index] << (index % 8));
        return result;
    }

    private static void AppendBits(ICollection<int> bits, int value, int count) {
        for (int index = 0; index < count; index++) bits.Add((value >> index) & 1);
    }

    private static void AppendCode(ICollection<int> bits, int value, int count) {
        for (int index = count - 1; index >= 0; index--) bits.Add((value >> index) & 1);
    }

    private static uint Checksum(byte[] data, int offset, int count) {
        uint value = 0;
        int end = offset + count;
        while (offset <= end - 4) {
            value ^= (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
            offset += 4;
        }
        uint tail = 0;
        while (offset < end) tail = (tail << 8) | data[offset++];
        return value ^ tail;
    }

    private static void UInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)value; data[offset + 1] = (byte)(value >> 8);
    }

    private static void UInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value; data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16); data[offset + 3] = (byte)(value >> 24);
    }

    private sealed class RejectReadStream : MemoryStream {
        internal RejectReadStream(byte[] data) : base(data, writable: false) { }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("The limit must be checked before reading.");
    }

    private sealed class CancelAfterReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal CancelAfterReadStream(byte[] data, CancellationTokenSource cancellation) : base(data, writable: false) { _cancellation = cancellation; }
        public override int Read(byte[] buffer, int offset, int count) {
            int result = base.Read(buffer, offset, Math.Min(count, 1));
            _cancellation.Cancel();
            return result;
        }
    }
}
