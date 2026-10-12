using System.Buffers.Binary;
using System.IO.Compression;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Counts the emitted string-storage records outside timed writer operations.</summary>
    internal static class XlsbStringStorageValidation {
        // MS-XLSB 2.1.4 framing and 2.3.2 record enumeration. This bounded inspection
        // does not decode workbook values; independent readers validate those separately.
        internal static void Validate(ZipArchive package, int rows, string engine, bool? sharedStrings) {
            ZipArchiveEntry? table = package.GetEntry("xl/sharedStrings.bin");
            int items = 0;
            uint? total = null, unique = null;
            if (table != null) Visit(Read(table), (type, payload) => {
                if (type == 159) { // BrtBeginSst
                    if (payload.Length != 8 || total != null) throw new InvalidDataException("Unexpected SST header.");
                    total = BinaryPrimitives.ReadUInt32LittleEndian(payload.Span);
                    unique = BinaryPrimitives.ReadUInt32LittleEndian(payload.Span[4..]);
                } else if (type == 19) items++; // BrtSSTItem
            });
            int inline = 0, references = 0;
            Visit(Read(package.GetEntry("xl/worksheets/sheet1.bin")!), (type, payload) => {
                if (type is 6 or 62) inline++; // BrtCellSt/BrtCellRString: string content in the cell record.
                if (type == 7) { // BrtCellIsst: index into the shared-string table.
                    if (payload.Length != 12 || BinaryPrimitives.ReadUInt32LittleEndian(payload.Span[8..]) >= items)
                        throw new InvalidDataException("Shared-string reference is invalid.");
                    references++;
                }
            });
            int textCells = checked(rows + 4);
            if (inline + references != textCells || (table != null && (total == null || unique != items)))
                throw new InvalidDataException("XLSB text-cell/table counts differ from the complete fixture.");
            if (sharedStrings == true && (inline != 0 || references != textCells || total != textCells
                || items != Math.Min(rows, 8) + 4)) throw new InvalidDataException("Requested XLSB shared-string policy differs.");
            if (sharedStrings == false && (inline != textCells || references != 0))
                throw new InvalidDataException("Requested XLSB inline-string policy differs.");
            Console.WriteLine($"XLSB storage={engine}; configured={(sharedStrings?.ToString() ?? "default")}; "
                + $"inlineCells={inline}; sharedStringCells={references}; hasSstPart={table != null}; "
                + $"sstReferences={total?.ToString() ?? "absent"}; sstUnique={items}.");
        }
        private static byte[] Read(ZipArchiveEntry entry) {
            using Stream source = entry.Open();
            using MemoryStream output = new MemoryStream();
            source.CopyTo(output);
            return output.ToArray();
        }
        private static void Visit(byte[] bytes, Action<int, ReadOnlyMemory<byte>> visit) {
            int position = 0;
            int Take() => position < bytes.Length ? bytes[position++]
                : throw new InvalidDataException("Truncated XLSB record header.");
            while (position < bytes.Length) {
                int first = Take();
                int type = first & 127;
                if ((first & 128) != 0) type |= (Take() & 127) << 7;
                int size = 0;
                for (int index = 0; index < 4; index++) {
                    int value = Take();
                    size |= (value & 127) << (index * 7);
                    if ((value & 128) == 0 || index == 3) break;
                }
                if (size > bytes.Length - position) throw new InvalidDataException("Truncated XLSB record payload.");
                visit(type, bytes.AsMemory(position, size));
                position += size;
            }
        }
    }
}
