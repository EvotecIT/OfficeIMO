using System.Text;

namespace OfficeIMO.OneNote;

internal static class OneNoteCabinetArchiveWriter {
    private const int BlockSize = 32 * 1024;
    private const ushort UtfNameAttribute = 0x0080;

    internal static byte[] Write(IReadOnlyList<OneNoteCabinetEntry> entries, long maxOutputBytes) {
        if (entries == null) throw new ArgumentNullException(nameof(entries));
        if (entries.Count == 0 || entries.Count > ushort.MaxValue) throw new OneNoteFormatException("ONENOTE_CAB_ENTRY_COUNT", "A OneNote package must contain between 1 and 65535 entries.");
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var folderData = new MemoryStream())
        using (var fileTable = new MemoryStream()) {
            foreach (OneNoteCabinetEntry entry in entries) {
                string name = OneNotePackageReader.NormalizeEntryName(entry.Name);
                if (!names.Add(name)) throw new OneNoteFormatException("ONENOTE_PACKAGE_DUPLICATE_ENTRY", "The .onepkg archive contains duplicate entry paths.");
                if (entry.Data.LongLength > uint.MaxValue || folderData.Length > uint.MaxValue - entry.Data.LongLength) {
                    throw new IOException("The uncompressed OneNote package exceeds the Cabinet folder limit.");
                }
                byte[] encodedName = Encoding.UTF8.GetBytes(name.Replace('/', '\\'));
                WriteUInt32(fileTable, (uint)entry.Data.Length);
                WriteUInt32(fileTable, (uint)folderData.Length);
                WriteUInt16(fileTable, 0);
                WriteUInt16(fileTable, 0);
                WriteUInt16(fileTable, 0);
                WriteUInt16(fileTable, UtfNameAttribute);
                fileTable.Write(encodedName, 0, encodedName.Length);
                fileTable.WriteByte(0);
                folderData.Write(entry.Data, 0, entry.Data.Length);
            }

            int blockCount = checked((int)((folderData.Length + BlockSize - 1) / BlockSize));
            if (blockCount > ushort.MaxValue) throw new IOException("The uncompressed OneNote package requires too many Cabinet data blocks.");
            const uint fileTableOffset = 44;
            long dataOffset = fileTableOffset + fileTable.Length;
            long cabinetSize = checked(dataOffset + folderData.Length + blockCount * 8L);
            int outputCapacity = GetOutputCapacity(cabinetSize, maxOutputBytes);

            using (var output = new MemoryStream(outputCapacity)) {
                output.WriteByte((byte)'M'); output.WriteByte((byte)'S'); output.WriteByte((byte)'C'); output.WriteByte((byte)'F');
                WriteUInt32(output, 0);
                WriteUInt32(output, (uint)cabinetSize);
                WriteUInt32(output, 0);
                WriteUInt32(output, fileTableOffset);
                WriteUInt32(output, 0);
                output.WriteByte(3);
                output.WriteByte(1);
                WriteUInt16(output, 1);
                WriteUInt16(output, (ushort)entries.Count);
                WriteUInt16(output, 0);
                WriteUInt16(output, 0);
                WriteUInt16(output, 0);
                WriteUInt32(output, (uint)dataOffset);
                WriteUInt16(output, (ushort)blockCount);
                WriteUInt16(output, 0);
                byte[] files = fileTable.ToArray();
                output.Write(files, 0, files.Length);
                byte[] data = folderData.ToArray();
                for (int offset = 0; offset < data.Length; offset += BlockSize) {
                    int count = Math.Min(BlockSize, data.Length - offset);
                    uint checksum = (uint)(count | (count << 16)) ^ OneNoteCabinetChecksum.Compute(data, offset, count);
                    WriteUInt32(output, checksum);
                    WriteUInt16(output, (ushort)count);
                    WriteUInt16(output, (ushort)count);
                    output.Write(data, offset, count);
                }
                return output.ToArray();
            }
        }
    }

    /// <summary>Validates the CAB size fields against both the caller limit and managed buffer capacity.</summary>
    internal static int GetOutputCapacity(long cabinetSize, long maxOutputBytes) {
        if (cabinetSize > uint.MaxValue || cabinetSize > maxOutputBytes) {
            throw new IOException("OneNote package output exceeds MaxOutputBytes.");
        }
        if (cabinetSize > int.MaxValue) {
            throw new IOException("OneNote package output exceeds the supported in-memory size.");
        }
        return checked((int)cabinetSize);
    }

    private static void WriteUInt16(Stream stream, ushort value) => FssHttpStreamObjectWriter.WriteUInt16(stream, value);
    private static void WriteUInt32(Stream stream, uint value) => FssHttpStreamObjectWriter.WriteUInt32(stream, value);
}
