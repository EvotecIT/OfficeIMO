using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    private static byte[] BuildXrefStreamEntries(Dictionary<int, int> offsets, int xrefObjectNumber) {
        using var stream = new MemoryStream();
        WriteXrefEntry(stream, 0, 0, 65535);
        for (int objectNumber = 1; objectNumber <= 8; objectNumber++) {
            if (objectNumber == xrefObjectNumber) {
                WriteXrefEntry(stream, 1, offsets[xrefObjectNumber], 0);
            } else if (offsets.TryGetValue(objectNumber, out int offset)) {
                WriteXrefEntry(stream, 1, offset, 0);
            } else {
                WriteXrefEntry(stream, 0, 0, 65535);
            }
        }

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamEntries(IReadOnlyDictionary<int, (int Type, int Field1, int Field2)> entries, int size) {
        using var stream = new MemoryStream();
        for (int objectNumber = 0; objectNumber < size; objectNumber++) {
            if (entries.TryGetValue(objectNumber, out var entry)) {
                WriteXrefEntry(stream, entry.Type, entry.Field1, entry.Field2);
            } else {
                WriteXrefEntry(stream, 0, 0, 65535);
            }
        }

        return stream.ToArray();
    }

    private static byte[] BuildXrefStreamEntries(IReadOnlyList<int> objectNumbers, IReadOnlyDictionary<int, (int Type, int Field1, int Field2)> entries) {
        using var stream = new MemoryStream();
        foreach (int objectNumber in objectNumbers) {
            if (!entries.TryGetValue(objectNumber, out var entry)) {
                throw new InvalidOperationException("Missing xref stream entry for object " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
            }

            WriteXrefEntry(stream, entry.Type, entry.Field1, entry.Field2);
        }

        return stream.ToArray();
    }

    private static void WriteClassicXrefTable(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int rootObjectNumber, int? previousXrefOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) + " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R");
        if (previousXrefOffset.HasValue) {
            WriteAscii(stream, " /Prev " + previousXrefOffset.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        WriteAscii(stream, " >>\n");
    }

    private static void WriteClassicXrefTableWithoutRoot(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int? previousXrefOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (previousXrefOffset.HasValue) {
            WriteAscii(stream, " /Prev " + previousXrefOffset.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        WriteAscii(stream, " >>\n");
    }

    private static void WriteClassicXrefTableWithXRefStm(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int rootObjectNumber, int xrefStreamOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /Root " + rootObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 R /XRefStm " + xrefStreamOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\n");
    }

    private static void WriteClassicXrefTableWithXRefStmWithoutRoot(Stream stream, IReadOnlyDictionary<int, int> entries, int size, int xrefStreamOffset) {
        WriteAscii(stream, "xref\n");
        var objectNumbers = entries.Keys.OrderBy(static objectNumber => objectNumber).ToList();
        int index = 0;
        while (index < objectNumbers.Count) {
            int first = objectNumbers[index];
            int end = index + 1;
            while (end < objectNumbers.Count && objectNumbers[end] == objectNumbers[end - 1] + 1) {
                end++;
            }

            WriteAscii(stream, first.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + (end - index).ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
            for (int i = index; i < end; i++) {
                int objectNumber = objectNumbers[i];
                if (objectNumber == 0) {
                    WriteAscii(stream, "0000000000 65535 f \n");
                } else {
                    WriteAscii(stream, entries[objectNumber].ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n");
                }
            }

            index = end;
        }

        WriteAscii(stream, "trailer\n<< /Size " + size.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /XRefStm " + xrefStreamOffset.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\n");
    }

    private static void WriteXrefEntry(Stream stream, int type, int field1, int field2) {
        stream.WriteByte((byte)type);
        WriteBigEndian(stream, field1, 4);
        WriteBigEndian(stream, field2, 2);
    }

    private static void WriteBigEndian(Stream stream, int value, int byteCount) {
        for (int shift = (byteCount - 1) * 8; shift >= 0; shift -= 8) {
            stream.WriteByte((byte)((value >> shift) & 0xFF));
        }
    }

    private static void WriteObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, string body) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n" + body + "\nendobj\n");
    }

    private static void WriteObjectGeneration(Stream stream, Dictionary<int, int> offsets, int objectNumber, int generation, string body) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " " + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " obj\n" + body + "\nendobj\n");
    }

    private static void WriteStreamObjectGeneration(Stream stream, Dictionary<int, int> offsets, int objectNumber, int generation, byte[] streamBytes) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " " + generation.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " obj\n<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        stream.Write(streamBytes, 0, streamBytes.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
    }

    private static void WriteStreamObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, byte[] streamBytes) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " 0 obj\n<< /Length " +
            streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " >>\nstream\n");
        stream.Write(streamBytes, 0, streamBytes.Length);
        WriteAscii(stream, "\nendstream\nendobj\n");
    }

    private static void WriteRawObject(Stream stream, Dictionary<int, int> offsets, int objectNumber, string objectText) {
        offsets[objectNumber] = (int)stream.Position;
        WriteAscii(stream, objectText);
        if (!objectText.EndsWith("\n", StringComparison.Ordinal)) {
            WriteAscii(stream, "\n");
        }
    }

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
