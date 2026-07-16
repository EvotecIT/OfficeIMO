namespace OfficeIMO.OneNote;

internal static class OneNoteRevisionStoreWriter {
    private const int HeaderLength = OneNoteFormatConstants.RevisionStoreHeaderLength;
    private const uint OneNote2010Build = 0x0FA129B4U;

    internal static byte[] Write(OneNoteWriteGraph graph, long maxOutputBytes = long.MaxValue) {
        if (maxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), "MaxOutputBytes must be greater than zero.");
        OneNoteDesktopWritePlan plan = OneNoteDesktopWritePlan.Create(graph);
        ulong position = HeaderLength;
        foreach (OneNoteDesktopFileNodeList list in plan.Lists) {
            position = Align(position);
            list.Offset = position;
            position = checked(position + list.Length);
        }
        foreach (OneNoteDesktopDataChunk data in plan.Data) {
            position = Align(position);
            data.Offset = position;
            position = checked(position + data.Length);
        }

        position = Align(position);
        ulong transactionOffset = position;
        int transactionEntriesLength = checked((plan.Lists.Count + 1) * 8);
        uint transactionLength = checked((uint)OneNoteDesktopBinary.Align8(transactionEntriesLength + 12L));
        ulong expectedLength = checked(transactionOffset + transactionLength);
        if (expectedLength > int.MaxValue) throw new IOException("Desktop OneNote output exceeds the supported in-memory size.");
        if (expectedLength > (ulong)maxOutputBytes) throw new IOException("Desktop OneNote output exceeds MaxOutputBytes.");

        var output = new byte[(int)expectedLength];
        WriteHeader(output, plan, transactionOffset, transactionLength, expectedLength);
        foreach (OneNoteDesktopFileNodeList list in plan.Lists) Copy(output, list.Offset, list.Encode());
        foreach (OneNoteDesktopDataChunk data in plan.Data) Copy(output, data.Offset, data.Data);
        WriteTransactionLog(output, transactionOffset, transactionLength, plan.Lists, graph.FileKind);
        return output;
    }

    private static void WriteHeader(
        byte[] output,
        OneNoteDesktopWritePlan plan,
        ulong transactionOffset,
        uint transactionLength,
        ulong expectedLength) {
        OneNoteWriteGraph graph = plan.Graph;
        uint version = graph.FileKind == OneNoteFileKind.Section ? 0x2AU : 0x1BU;
        using (var stream = new MemoryStream(output, true)) {
            OneNoteDesktopBinary.WriteGuid(stream, graph.FileKind == OneNoteFileKind.Section
                ? OneNoteFormatConstants.SectionFileType
                : OneNoteFormatConstants.TableOfContentsFileType);
            OneNoteDesktopBinary.WriteGuid(stream, graph.FileId);
            OneNoteDesktopBinary.WriteGuid(stream, Guid.Empty);
            OneNoteDesktopBinary.WriteGuid(stream, OneNoteFormatConstants.RevisionStoreFormat);
            for (int index = 0; index < 4; index++) FssHttpStreamObjectWriter.WriteUInt32(stream, version);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, uint.MaxValue);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 1);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt64(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, uint.MaxValue);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            OneNoteDesktopBinary.WriteGuid(stream, graph.AncestorId);
            FssHttpStreamObjectWriter.WriteUInt32(stream, graph.FileNameCrc);
            OneNoteDesktopBinary.WriteNilReference(stream);
            OneNoteDesktopBinary.WriteReference(stream, transactionOffset, transactionLength);
            OneNoteDesktopBinary.WriteReference(stream, plan.Root.Offset, plan.Root.Length);
            OneNoteDesktopBinary.WriteNilReference(stream);
            FssHttpStreamObjectWriter.WriteUInt64(stream, expectedLength);
            FssHttpStreamObjectWriter.WriteUInt64(stream, 0);
            OneNoteDesktopBinary.WriteGuid(stream, Guid.NewGuid());
            FssHttpStreamObjectWriter.WriteUInt64(stream, 1);
            OneNoteDesktopBinary.WriteGuid(stream, Guid.NewGuid());
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            OneNoteDesktopBinary.WriteReference(stream, 0, 0);
            OneNoteDesktopBinary.WriteReference(stream, 0, 0);
            for (int index = 0; index < 4; index++) FssHttpStreamObjectWriter.WriteUInt32(stream, OneNote2010Build);
        }
    }

    private static void WriteTransactionLog(
        byte[] output,
        ulong offset,
        uint length,
        IReadOnlyList<OneNoteDesktopFileNodeList> lists,
        OneNoteFileKind fileKind) {
        using (var stream = new MemoryStream(output, true)) {
            stream.Position = checked((long)offset);
            foreach (OneNoteDesktopFileNodeList list in lists) {
                FssHttpStreamObjectWriter.WriteUInt32(stream, list.Id);
                FssHttpStreamObjectWriter.WriteUInt32(stream, checked((uint)list.Nodes.Count));
            }
            int transactionBytes = checked(lists.Count * 8);
            uint crc = OneNoteCrc32.Continue(
                0,
                output,
                checked((int)offset),
                transactionBytes,
                fileKind);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 1);
            FssHttpStreamObjectWriter.WriteUInt32(stream, crc);
            OneNoteDesktopBinary.WriteNilReference(stream);
            long end = checked((long)(offset + length));
            if (stream.Position > end) throw new OneNoteFormatException("ONENOTE_WRITE_TRANSACTION_LOG", "The desktop OneNote transaction log exceeded its allocated fragment.");
        }
    }

    private static void Copy(byte[] output, ulong offset, byte[] data) {
        if (offset > int.MaxValue || data.Length > output.Length - (int)offset) throw new IOException("Desktop OneNote output layout exceeded the allocated buffer.");
        Buffer.BlockCopy(data, 0, output, (int)offset, data.Length);
    }

    private static ulong Align(ulong value) => checked((value + 7UL) & ~7UL);
}
