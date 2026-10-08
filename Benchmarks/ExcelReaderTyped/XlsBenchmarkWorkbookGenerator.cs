using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

// Upstream BIFF8/OLE fixture at ca5b50f99e8ef57ab476f0a2bc8043558d58b28d.
// Copyright (c) 2026 Gabriel Matte; MIT license in THIRD-PARTY-NOTICES.md.
    internal static class XlsBenchmarkWorkbookGenerator
    {
        private const int SectorSize = 512;
        private const int FatSector = unchecked((int)0xFFFFFFFD);
        private const int EndOfChain = unchecked((int)0xFFFFFFFE);
        private const int FreeSector = unchecked((int)0xFFFFFFFF);
        private static readonly string[] Pool =
            ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];

        internal static byte[] Build(int rows)
        {
            byte[] sheet = BuildSheet(rows);
            const int globalsLength = 20 + 24 + 24 + 14 + 4;
            using MemoryStream workbook = new();
            WriteRecord(workbook, 0x0809, [.. U16(0x0600), .. U16(0x0005), .. new byte[12]]);
            WriteRecord(workbook, 0x00E0, Xf(0));
            WriteRecord(workbook, 0x00E0, Xf(14));
            WriteRecord(workbook, 0x0085, [.. I32(globalsLength), 0, 0, 2, 0, (byte)'S', (byte)'1']);
            WriteRecord(workbook, 0x000A, []);
            if (workbook.Length != globalsLength)
            {
                throw new InvalidOperationException(
                    $"BoundSheet offset {globalsLength} must equal globals length {workbook.Length}.");
            }
            workbook.Write(sheet);
            return BuildOle(workbook.ToArray()).ToArray();
        }

        private static byte[] BuildSheet(int rows)
        {
            using MemoryStream sheet = new();
            WriteRecord(sheet, 0x0809, [.. U16(0x0600), .. U16(0x0010), .. new byte[12]]);
            WriteRecord(sheet, 0x0200, [.. I32(0), .. I32(rows), .. U16(0), .. U16(4), 0, 0]);
            for (int r = 1; r <= rows; r++)
            {
                WriteRecord(sheet, 0x0204, [.. U16(r - 1), .. U16(0), .. U16(0), .. BiffString(Pool[r % Pool.Length])]);
                WriteNumber(sheet, r - 1, 1, 0, r);
                WriteNumber(sheet, r - 1, 2, 1, 45292 + (r % 3650) + 0.25);
                WriteNumber(sheet, r - 1, 3, 0, r * 1.5);
            }
            WriteRecord(sheet, 0x000A, []);
            return sheet.ToArray();
        }

        private static MemoryStream BuildOle(byte[] workbook)
        {
            // Regular FAT streams require the OLE mini-stream cutoff size. This
            // only pads diagnostic inputs; the original 50,000 rows are unchanged.
            int workbookSize = Math.Max(4096, RoundUp(workbook.Length, SectorSize));
            int workbookSectorCount = workbookSize / SectorSize;
            int fatSectorCount = (workbookSectorCount + 1 + 126) / 127;
            int dirSector = fatSectorCount;
            int workbookStart = dirSector + 1;
            int totalSectors = fatSectorCount + 1 + workbookSectorCount;
            int fatEntries = fatSectorCount * 128;
            byte[] fat = [.. Enumerable.Repeat(FreeSector, fatEntries).SelectMany(I32)];
            for (int i = 0; i < fatSectorCount; i++)
            {
                WriteI32(fat, i * 4, FatSector);
            }
            WriteI32(fat, dirSector * 4, EndOfChain);
            for (int i = 0; i < workbookSectorCount; i++)
            {
                int sector = workbookStart + i;
                int next = i == workbookSectorCount - 1 ? EndOfChain : sector + 1;
                WriteI32(fat, sector * 4, next);
            }

            MemoryStream ole = new();
            ole.Write(Header(fatSectorCount, dirSector));
            ole.Write(fat);
            ole.Write(Directory(workbookStart, workbookSize));
            ole.Write(workbook);
            ole.Write(new byte[workbookSize - workbook.Length]);
            ole.Position = 0;
            _ = totalSectors;
            return ole;
        }

        private static byte[] Header(int fatSectorCount, int dirSector)
        {
            byte[] header = new byte[SectorSize];
            byte[] sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
            sig.CopyTo(header, 0);
            WriteU16(header, 0x18, 0x003E);
            WriteU16(header, 0x1A, 0x0003);
            WriteU16(header, 0x1C, 0xFFFE);
            WriteU16(header, 0x1E, 9);
            WriteU16(header, 0x20, 6);
            WriteI32(header, 0x2C, fatSectorCount);
            WriteI32(header, 0x30, dirSector);
            WriteI32(header, 0x38, 4096);
            WriteI32(header, 0x3C, EndOfChain);
            WriteI32(header, 0x40, 0);
            WriteI32(header, 0x44, EndOfChain);
            WriteI32(header, 0x48, 0);
            for (int i = 0x4C; i < SectorSize; i += 4)
            {
                WriteI32(header, i, FreeSector);
            }
            for (int i = 0; i < fatSectorCount; i++)
            {
                WriteI32(header, 0x4C + i * 4, i);
            }
            return header;
        }

        private static byte[] Directory(int workbookStart, int workbookSize)
        {
            byte[] dir = new byte[SectorSize];
            WriteDirectoryEntry(dir.AsSpan(0, 128), "Root Entry", 5, EndOfChain, 0, child: 1);
            WriteDirectoryEntry(dir.AsSpan(128, 128), "Workbook", 2, workbookStart, workbookSize, child: EndOfChain);
            return dir;
        }

        private static void WriteDirectoryEntry(Span<byte> entry, string name, byte type, int startSector, long size, int child)
        {
            Encoding.Unicode.GetBytes(name + '\0').CopyTo(entry);
            WriteU16(entry, 64, (ushort)((name.Length + 1) * 2));
            entry[66] = type;
            entry[67] = 1;
            WriteI32(entry, 68, EndOfChain);
            WriteI32(entry, 72, EndOfChain);
            WriteI32(entry, 76, child);
            WriteI32(entry, 116, startSector);
            BinaryPrimitives.WriteInt64LittleEndian(entry[120..], size);
        }

        private static void WriteNumber(MemoryStream sheet, int row, int col, int xf, double value)
        {
            WriteRecord(sheet, 0x0203, [.. U16(row), .. U16(col), .. U16(xf), .. Double(value)]);
        }

        private static byte[] Xf(int format)
        {
            byte[] data = new byte[20];
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(2), (ushort)format);
            return data;
        }

        private static byte[] BiffString(string value)
        {
            byte[] text = Encoding.Latin1.GetBytes(value);
            return [.. U16(value.Length), 0, .. text];
        }

        private static void WriteRecord(MemoryStream stream, int id, byte[] data)
        {
            stream.Write(U16(id));
            stream.Write(U16(data.Length));
            stream.Write(data);
        }

        private static int RoundUp(int value, int multiple)
        {
            return (value + multiple - 1) / multiple * multiple;
        }

        private static byte[] U16(int value)
        {
            byte[] bytes = new byte[2];
            WriteU16(bytes, 0, (ushort)value);
            return bytes;
        }

        private static byte[] I32(int value)
        {
            byte[] bytes = new byte[4];
            WriteI32(bytes, 0, value);
            return bytes;
        }

        private static byte[] Double(double value)
        {
            byte[] bytes = new byte[8];
            BinaryPrimitives.WriteDoubleLittleEndian(bytes, value);
            return bytes;
        }

        private static void WriteU16(Span<byte> bytes, int offset, ushort value)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(bytes.Slice(offset, 2), value);
        }

        private static void WriteI32(Span<byte> bytes, int offset, int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(bytes.Slice(offset, 4), value);
        }
    }
