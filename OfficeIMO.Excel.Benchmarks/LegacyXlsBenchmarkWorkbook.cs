using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Builds BIFF8 workbooks with the string-cell records that OfficeIMO's writer does not emit
/// (inline LABEL cells, cached formula STRING results and shared strings continued across
/// CONTINUE records) and wraps them in a minimal version 3 compound file.
/// </summary>
internal static class LegacyXlsBenchmarkWorkbook {
    internal const string SheetName = "Strings";
    internal const int ContinuedSharedStringCount = 512;
    internal const int ContinuedSharedStringLength = 6000;
    private const int MaximumRecordPayload = 8224;
    private const byte HighByteOption = 0x01;

    internal static byte[] Create(string cells, int rowCount, int columnCount) {
        ArgumentOutOfRangeException.ThrowIfLessThan(rowCount, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(columnCount, 1);
        if (rowCount > ushort.MaxValue + 1 || columnCount > 256) {
            throw new ArgumentOutOfRangeException(nameof(rowCount), "BIFF8 worksheets hold 65,536 rows and 256 columns.");
        }

        using var stream = new MemoryStream();
        WriteRecord(stream, 0x0809, [0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07]);
        long boundSheetPosition = stream.Position;
        WriteRecord(stream, 0x0085, BuildBoundSheetPayload(SheetName));
        if (cells == "ContinuedSharedString") {
            WriteSharedStrings(stream, rowCount * columnCount);
        }
        WriteRecord(stream, 0x000a, []);

        int sheetOffset = checked((int)stream.Position);
        WriteRecord(stream, 0x0809, [0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07]);
        for (int row = 0; row < rowCount; row++) {
            for (int column = 0; column < columnCount; column++) {
                WriteCell(stream, cells, row, column, columnCount);
            }
        }
        WriteRecord(stream, 0x000a, []);

        byte[] workbookStream = stream.ToArray();
        BinaryPrimitives.WriteInt32LittleEndian(workbookStream.AsSpan(checked((int)boundSheetPosition + 4)), sheetOffset);
        return WrapInCompoundFile(workbookStream);
    }

    /// <summary>The text each cell must deliver.</summary>
    internal static string Text(string cells, int row, int column, int columnCount) {
        return cells switch {
            "Label" or "FormulaString" => $"Café {row:D6}-{column} détail Ø",
            "LabelUtf16" => $"Zażółć {row:D6}-{column} Ωmega",
            "ContinuedSharedString" => SharedString((row * columnCount + column) % ContinuedSharedStringCount),
            _ => throw new ArgumentOutOfRangeException(nameof(cells), cells, "Unknown XLS string cell shape.")
        };
    }

    private static readonly string[] SharedStrings = CreateSharedStrings();

    private static string SharedString(int index) => SharedStrings[index];

    private static string[] CreateSharedStrings() {
        const string pattern = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ";
        var strings = new string[ContinuedSharedStringCount];
        var builder = new StringBuilder(ContinuedSharedStringLength);
        for (int index = 0; index < strings.Length; index++) {
            builder.Clear().Append("S").Append(index.ToString("D4", System.Globalization.CultureInfo.InvariantCulture)).Append(' ');
            while (builder.Length < ContinuedSharedStringLength) {
                builder.Append(pattern, 0, Math.Min(pattern.Length, ContinuedSharedStringLength - builder.Length));
            }
            strings[index] = builder.ToString();
        }
        return strings;
    }

    private static void WriteCell(Stream stream, string cells, int row, int column, int columnCount) {
        string text = Text(cells, row, column, columnCount);
        switch (cells) {
            case "Label":
                WriteRecord(stream, 0x0204, BuildCellPayload(row, column, payload => WriteBiffString(payload, text, utf16: false)));
                break;
            case "LabelUtf16":
                WriteRecord(stream, 0x0204, BuildCellPayload(row, column, payload => WriteBiffString(payload, text, utf16: true)));
                break;
            case "FormulaString": {
                // A FORMULA whose cached result is a string, followed by its STRING record.
                byte[] formula = new byte[22];
                BinaryPrimitives.WriteUInt16LittleEndian(formula, (ushort)row);
                BinaryPrimitives.WriteUInt16LittleEndian(formula.AsSpan(2), (ushort)column);
                BinaryPrimitives.WriteUInt16LittleEndian(formula.AsSpan(12), 0xffff);
                WriteRecord(stream, 0x0006, formula);
                using var payload = new MemoryStream();
                WriteBiffString(payload, text, utf16: false);
                WriteRecord(stream, 0x0207, payload.ToArray());
                break;
            }
            case "ContinuedSharedString": {
                byte[] payload = new byte[10];
                BinaryPrimitives.WriteUInt16LittleEndian(payload, (ushort)row);
                BinaryPrimitives.WriteUInt16LittleEndian(payload.AsSpan(2), (ushort)column);
                BinaryPrimitives.WriteUInt32LittleEndian(payload.AsSpan(6), (uint)((row * columnCount + column) % ContinuedSharedStringCount));
                WriteRecord(stream, 0x00fd, payload);
                break;
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(cells), cells, "Unknown XLS string cell shape.");
        }
    }

    /// <summary>
    /// Writes an SST whose compressed strings continue into CONTINUE records. A continued
    /// string's characters resume after the CONTINUE record's option byte.
    /// </summary>
    private static void WriteSharedStrings(Stream stream, int totalReferences) {
        var records = new List<byte[]>();
        var current = new MemoryStream();
        Span<byte> header = stackalloc byte[8];
        BinaryPrimitives.WriteUInt32LittleEndian(header, (uint)totalReferences);
        BinaryPrimitives.WriteUInt32LittleEndian(header.Slice(4), ContinuedSharedStringCount);
        current.Write(header);
        foreach (string text in SharedStrings) {
            if (MaximumRecordPayload - current.Length < 4) {
                records.Add(current.ToArray());
                current = new MemoryStream();
            }
            Span<byte> stringHeader = stackalloc byte[3];
            BinaryPrimitives.WriteUInt16LittleEndian(stringHeader, (ushort)text.Length);
            stringHeader[2] = 0;
            current.Write(stringHeader);
            int written = 0;
            while (written < text.Length) {
                int available = checked((int)(MaximumRecordPayload - current.Length));
                if (available == 0) {
                    records.Add(current.ToArray());
                    current = new MemoryStream();
                    current.WriteByte(0);
                    continue;
                }
                int take = Math.Min(available, text.Length - written);
                for (int index = 0; index < take; index++) {
                    current.WriteByte(checked((byte)text[written + index]));
                }
                written += take;
            }
        }
        records.Add(current.ToArray());
        if (records.Count < 2) throw new InvalidOperationException("The shared-string fixture must continue across records.");
        for (int index = 0; index < records.Count; index++) {
            WriteRecord(stream, index == 0 ? (ushort)0x00fc : (ushort)0x003c, records[index]);
        }
    }

    private static byte[] BuildCellPayload(int row, int column, Action<Stream> writeValue) {
        using var payload = new MemoryStream();
        Span<byte> cell = stackalloc byte[6];
        BinaryPrimitives.WriteUInt16LittleEndian(cell, (ushort)row);
        BinaryPrimitives.WriteUInt16LittleEndian(cell.Slice(2), (ushort)column);
        BinaryPrimitives.WriteUInt16LittleEndian(cell.Slice(4), 0);
        payload.Write(cell);
        writeValue(payload);
        return payload.ToArray();
    }

    private static void WriteBiffString(Stream stream, string text, bool utf16) {
        Span<byte> header = stackalloc byte[3];
        BinaryPrimitives.WriteUInt16LittleEndian(header, checked((ushort)text.Length));
        header[2] = utf16 ? HighByteOption : (byte)0;
        stream.Write(header);
        if (utf16) {
            stream.Write(Encoding.Unicode.GetBytes(text));
            return;
        }
        foreach (char character in text) {
            stream.WriteByte(checked((byte)character));
        }
    }

    private static byte[] BuildBoundSheetPayload(string name) {
        byte[] payload = new byte[8 + name.Length];
        payload[6] = (byte)name.Length;
        Encoding.ASCII.GetBytes(name, 0, name.Length, payload, 8);
        return payload;
    }

    private static void WriteRecord(Stream stream, ushort type, byte[] payload) {
        if (payload.Length > MaximumRecordPayload) throw new InvalidOperationException("A BIFF8 record payload exceeds 8,224 bytes.");
        Span<byte> header = stackalloc byte[4];
        BinaryPrimitives.WriteUInt16LittleEndian(header, type);
        BinaryPrimitives.WriteUInt16LittleEndian(header.Slice(2), (ushort)payload.Length);
        stream.Write(header);
        stream.Write(payload);
    }

    /// <summary>
    /// Wraps a workbook stream as the only stream of a version 3 compound file. The stream
    /// must exceed the 4,096-byte mini-stream cutoff so it lives in regular sectors.
    /// </summary>
    private static byte[] WrapInCompoundFile(byte[] workbookStream) {
        const int sectorSize = 512;
        const uint freeSector = 0xffffffff;
        const uint endOfChain = 0xfffffffe;
        const uint fatSector = 0xfffffffd;
        const uint noStream = 0xffffffff;
        if (workbookStream.Length < 4096) throw new InvalidOperationException("The workbook stream must exceed the mini-stream cutoff.");

        int streamSectors = (workbookStream.Length + sectorSize - 1) / sectorSize;
        int fatSectors = 1;
        while (fatSectors * (sectorSize / 4) < streamSectors + 1 + fatSectors) fatSectors++;
        if (fatSectors > 109) throw new InvalidOperationException("The fixture exceeds the header's FAT sector table.");
        int directorySector = streamSectors;
        int firstFatSector = streamSectors + 1;
        int totalSectors = streamSectors + 1 + fatSectors;
        byte[] file = new byte[(totalSectors + 1) * sectorSize];
        Span<byte> headerBytes = file.AsSpan(0, sectorSize);

        new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }.CopyTo(headerBytes);
        BinaryPrimitives.WriteUInt16LittleEndian(headerBytes.Slice(0x18), 0x003e);
        BinaryPrimitives.WriteUInt16LittleEndian(headerBytes.Slice(0x1a), 0x0003);
        BinaryPrimitives.WriteUInt16LittleEndian(headerBytes.Slice(0x1c), 0xfffe);
        BinaryPrimitives.WriteUInt16LittleEndian(headerBytes.Slice(0x1e), 9);
        BinaryPrimitives.WriteUInt16LittleEndian(headerBytes.Slice(0x20), 6);
        BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x2c), (uint)fatSectors);
        BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x30), (uint)directorySector);
        BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x38), 4096);
        BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x3c), endOfChain);
        BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x44), endOfChain);
        for (int index = 0; index < 109; index++) {
            uint entry = index < fatSectors ? (uint)(firstFatSector + index) : freeSector;
            BinaryPrimitives.WriteUInt32LittleEndian(headerBytes.Slice(0x4c + index * 4), entry);
        }

        workbookStream.CopyTo(file.AsSpan(sectorSize));

        Span<byte> directory = file.AsSpan((directorySector + 1) * sectorSize, sectorSize);
        WriteDirectoryEntry(directory.Slice(0, 128), "Root Entry", type: 5, child: 1, startSector: endOfChain, size: 0);
        WriteDirectoryEntry(directory.Slice(128, 128), "Workbook", type: 2, child: noStream, startSector: 0, size: (uint)workbookStream.Length);
        for (int index = 2; index < 4; index++) {
            Span<byte> unused = directory.Slice(index * 128, 128);
            BinaryPrimitives.WriteUInt32LittleEndian(unused.Slice(0x44), noStream);
            BinaryPrimitives.WriteUInt32LittleEndian(unused.Slice(0x48), noStream);
            BinaryPrimitives.WriteUInt32LittleEndian(unused.Slice(0x4c), noStream);
        }

        Span<byte> fat = file.AsSpan((firstFatSector + 1) * sectorSize, fatSectors * sectorSize);
        fat.Fill(0xff);
        for (int sector = 0; sector < streamSectors; sector++) {
            uint next = sector == streamSectors - 1 ? endOfChain : (uint)(sector + 1);
            BinaryPrimitives.WriteUInt32LittleEndian(fat.Slice(sector * 4), next);
        }
        BinaryPrimitives.WriteUInt32LittleEndian(fat.Slice(directorySector * 4), endOfChain);
        for (int index = 0; index < fatSectors; index++) {
            BinaryPrimitives.WriteUInt32LittleEndian(fat.Slice((firstFatSector + index) * 4), fatSector);
        }
        return file;
    }

    private static void WriteDirectoryEntry(Span<byte> entry, string name, byte type, uint child, uint startSector, uint size) {
        const uint noStream = 0xffffffff;
        Encoding.Unicode.GetBytes(name).CopyTo(entry);
        BinaryPrimitives.WriteUInt16LittleEndian(entry.Slice(0x40), (ushort)((name.Length + 1) * 2));
        entry[0x42] = type;
        entry[0x43] = 1;
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(0x44), noStream);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(0x48), noStream);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(0x4c), child);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(0x74), startSector);
        BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(0x78), size);
    }
}
