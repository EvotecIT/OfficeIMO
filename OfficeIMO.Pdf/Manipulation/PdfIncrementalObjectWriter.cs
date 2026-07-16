using System.Globalization;

namespace OfficeIMO.Pdf;

internal enum PdfIncrementalXrefFormat {
    Automatic,
    ClassicTable,
    XrefStream
}

/// <summary>Shared serializer for append-only indirect-object revisions.</summary>
internal static class PdfIncrementalObjectWriter {
    public static byte[] Append(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        IEnumerable<int>? changedObjectNumbers = null,
        IReadOnlyList<(int ObjectNumber, byte[] Bytes)>? rawObjects = null,
        int? infoObjectNumberOverride = null,
        PdfIncrementalXrefFormat format = PdfIncrementalXrefFormat.Automatic,
        PdfStandardSecurityHandler? encryptionHandler = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(security, nameof(security));
        Guard.NotNull(trailerRaw, nameof(trailerRaw));

        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental update.");
        }

        if (security.HasEncryption && encryptionHandler is null) {
            throw new NotSupportedException("Appending objects in an encrypted PDF requires an authenticated encryption context.");
        }

        IReadOnlyList<(int ObjectNumber, byte[] Bytes)> effectiveRawObjects = rawObjects ?? Array.Empty<(int ObjectNumber, byte[] Bytes)>();
        if (security.HasEncryption && effectiveRawObjects.Count > 0) {
            throw new NotSupportedException("Raw incremental objects cannot be appended to an encrypted PDF. Supply typed PDF objects so strings and streams are encrypted with their object keys.");
        }

        List<SerializedObject> serialized = SerializeObjects(objects, changedObjectNumbers, effectiveRawObjects, encryptionHandler);
        if (serialized.Count == 0) {
            throw new ArgumentException("At least one changed or new indirect object is required for an incremental update.", nameof(changedObjectNumbers));
        }

        PdfIncrementalXrefFormat effectiveFormat = format == PdfIncrementalXrefFormat.Automatic
            ? security.HasXrefStreams ? PdfIncrementalXrefFormat.XrefStream : PdfIncrementalXrefFormat.ClassicTable
            : format;
        if (effectiveFormat != PdfIncrementalXrefFormat.ClassicTable && effectiveFormat != PdfIncrementalXrefFormat.XrefStream) {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported incremental cross-reference format.");
        }

        using var output = new MemoryStream(pdf.Length + serialized.Sum(static item => item.Bytes.Length) + (serialized.Count * 48) + 512);
        output.Write(pdf, 0, pdf.Length);
        EnsureLineBreak(output, pdf);

        var offsets = new Dictionary<int, long>();
        for (int i = 0; i < serialized.Count; i++) {
            SerializedObject item = serialized[i];
            offsets.Add(item.ObjectNumber, output.Position);
            output.Write(item.Bytes, 0, item.Bytes.Length);
        }

        int maximumObjectNumber = Math.Max(objects.Count == 0 ? 0 : objects.Keys.Max(), serialized.Max(static item => item.ObjectNumber));
        int infoObjectNumber = infoObjectNumberOverride ?? security.InfoObjectNumber ?? 0;
        if (effectiveFormat == PdfIncrementalXrefFormat.ClassicTable) {
            WriteClassicXref(output, objects, security, trailerRaw, serialized, offsets, maximumObjectNumber + 1, infoObjectNumber);
        } else {
            WriteXrefStream(output, objects, security, trailerRaw, serialized, offsets, maximumObjectNumber, infoObjectNumber);
        }

        return output.ToArray();
    }

    private static List<SerializedObject> SerializeObjects(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int>? changedObjectNumbers,
        IReadOnlyList<(int ObjectNumber, byte[] Bytes)> rawObjects,
        PdfStandardSecurityHandler? encryptionHandler) {
        var contextObjects = new Dictionary<int, PdfIndirectObject>(objects);
        for (int i = 0; i < rawObjects.Count; i++) {
            if (!contextObjects.ContainsKey(rawObjects[i].ObjectNumber)) {
                contextObjects.Add(rawObjects[i].ObjectNumber, new PdfIndirectObject(rawObjects[i].ObjectNumber, 0, PdfNull.Instance));
            }
        }

        var rawByObjectNumber = new Dictionary<int, byte[]>();
        for (int i = 0; i < rawObjects.Count; i++) {
            if (rawByObjectNumber.ContainsKey(rawObjects[i].ObjectNumber)) {
                throw new ArgumentException("Raw incremental objects must have unique object numbers.", nameof(rawObjects));
            }

            rawByObjectNumber.Add(rawObjects[i].ObjectNumber, rawObjects[i].Bytes);
        }

        int[] objectNumbers = (changedObjectNumbers ?? Array.Empty<int>())
            .Concat(rawByObjectNumber.Keys)
            .Distinct()
            .OrderBy(static objectNumber => objectNumber)
            .ToArray();
        var identityMap = contextObjects.Keys.ToDictionary(static objectNumber => objectNumber, static objectNumber => objectNumber);
        var context = new PdfPageExtractor.SerializationContext(
            identityMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            contextObjects,
            preserveReferenceGenerations: true,
            preserveRawStringBytes: encryptionHandler is not null);
        var serialized = new List<SerializedObject>(objectNumbers.Length);
        for (int i = 0; i < objectNumbers.Length; i++) {
            int objectNumber = objectNumbers[i];
            if (rawByObjectNumber.TryGetValue(objectNumber, out byte[]? rawBytes)) {
                serialized.Add(new SerializedObject(objectNumber, 0, rawBytes));
                continue;
            }

            if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) {
                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was changed but could not be found.");
            }

            PdfObject value = encryptionHandler is null
                ? indirect.Value
                : encryptionHandler.EncryptObject(objectNumber, indirect.Generation, indirect.Value);
            serialized.Add(new SerializedObject(
                objectNumber,
                indirect.Generation,
                PdfObjectBytes.WrapIndirectObject(
                    objectNumber,
                    indirect.Generation,
                    PdfPageExtractor.SerializeObject(value, context))));
        }

        return serialized;
    }

    private static void WriteClassicXref(
        Stream output,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        IReadOnlyList<SerializedObject> serialized,
        Dictionary<int, long> offsets,
        int size,
        int infoObjectNumber) {
        long xrefOffset = output.Position;
        using var writer = CreateWriter(output);
        writer.WriteLine("xref");
        for (int i = 0; i < serialized.Count; i++) {
            SerializedObject item = serialized[i];
            writer.WriteLine(item.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " 1");
            writer.WriteLine(offsets[item.ObjectNumber].ToString("0000000000", CultureInfo.InvariantCulture) + " " + item.Generation.ToString("00000", CultureInfo.InvariantCulture) + " n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine(BuildTrailerDictionary(objects, security, trailerRaw, size, infoObjectNumber));
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();
    }

    private static void WriteXrefStream(
        Stream output,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        IReadOnlyList<SerializedObject> serialized,
        Dictionary<int, long> offsets,
        int maximumObjectNumber,
        int infoObjectNumber) {
        int xrefObjectNumber = maximumObjectNumber + 1;
        long xrefOffset = output.Position;
        offsets.Add(xrefObjectNumber, xrefOffset);
        var entries = serialized
            .Select(static item => (item.ObjectNumber, item.Generation))
            .Append((ObjectNumber: xrefObjectNumber, Generation: 0))
            .OrderBy(static item => item.ObjectNumber)
            .ToArray();
        byte[] streamBytes = BuildXrefStreamEntries(entries, offsets);
        int size = xrefObjectNumber + 1;
        string index = string.Join(" ", entries.Select(static item => item.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " 1"));
        string trailerEntries = BuildTrailerEntries(objects, security, trailerRaw, infoObjectNumber);
        string dictionary = "<< /Type /XRef /Size " + size.ToString(CultureInfo.InvariantCulture) +
            " /W [1 8 2] /Index [" + index + "]" + trailerEntries +
            " /Length " + streamBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>";

        using (var writer = CreateWriter(output)) {
            writer.WriteLine(xrefObjectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj");
            writer.WriteLine(dictionary);
            writer.WriteLine("stream");
            writer.Flush();
        }

        output.Write(streamBytes, 0, streamBytes.Length);
        using (var writer = CreateWriter(output)) {
            writer.WriteLine();
            writer.WriteLine("endstream");
            writer.WriteLine("endobj");
            writer.WriteLine("startxref");
            writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
            writer.WriteLine("%%EOF");
            writer.Flush();
        }
    }

    private static byte[] BuildXrefStreamEntries(
        (int ObjectNumber, int Generation)[] entries,
        Dictionary<int, long> offsets) {
        var bytes = new byte[entries.Length * 11];
        for (int i = 0; i < entries.Length; i++) {
            int position = i * 11;
            bytes[position] = 1;
            WriteBigEndian(bytes, position + 1, offsets[entries[i].ObjectNumber], 8);
            WriteBigEndian(bytes, position + 9, entries[i].Generation, 2);
        }

        return bytes;
    }

    private static void WriteBigEndian(byte[] destination, int offset, long value, int length) {
        for (int i = length - 1; i >= 0; i--) {
            destination[offset + i] = (byte)(value & 0xFF);
            value >>= 8;
        }
    }

    private static string BuildTrailerDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        int size,
        int infoObjectNumber) =>
        "<< /Size " + size.ToString(CultureInfo.InvariantCulture) + BuildTrailerEntries(objects, security, trailerRaw, infoObjectNumber) + " >>";

    private static string BuildTrailerEntries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        int infoObjectNumber) =>
        " /Root " + BuildExistingReference(objects, security.RootObjectNumber!.Value) +
        (infoObjectNumber > 0 ? " /Info " + BuildExistingReference(objects, infoObjectNumber) : string.Empty) +
        (security.EncryptObjectNumber.HasValue ? " /Encrypt " + BuildExistingReference(objects, security.EncryptObjectNumber.Value) : string.Empty) +
        " /Prev " + security.LastStartXrefOffset!.Value.ToString(CultureInfo.InvariantCulture) +
        ReadTrailerIdEntry(trailerRaw);

    private static string BuildExistingReference(Dictionary<int, PdfIndirectObject> objects, int objectNumber) {
        int generation = objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ? indirect.Generation : 0;
        return PdfSyntaxEscaper.IndirectReference(objectNumber, generation);
    }

    internal static string ReadTrailerIdEntry(string trailerRaw) {
        int nameIndex = IndexOfName(trailerRaw, "ID");
        if (nameIndex < 0) {
            return string.Empty;
        }

        int start = trailerRaw.IndexOf('[', nameIndex);
        if (start < 0) {
            return string.Empty;
        }

        int depth = 0;
        for (int i = start; i < trailerRaw.Length; i++) {
            if (trailerRaw[i] == '[') {
                depth++;
            } else if (trailerRaw[i] == ']' && --depth == 0) {
                return " /ID " + trailerRaw.Substring(start, i - start + 1).Trim();
            }
        }

        return string.Empty;
    }

    private static int IndexOfName(string value, string name) {
        string token = "/" + name;
        int index = 0;
        while (index < value.Length) {
            int found = value.IndexOf(token, index, StringComparison.Ordinal);
            if (found < 0) {
                return -1;
            }

            int after = found + token.Length;
            if (after >= value.Length || IsDelimiter(value[after])) {
                return found;
            }

            index = after;
        }

        return -1;
    }

    private static bool IsDelimiter(char value) =>
        char.IsWhiteSpace(value) || value == '/' || value == '<' || value == '>' ||
        value == '[' || value == ']' || value == '(' || value == ')';

    private static void EnsureLineBreak(Stream output, byte[] pdf) {
        if (pdf.Length == 0 || (pdf[pdf.Length - 1] != (byte)'\n' && pdf[pdf.Length - 1] != (byte)'\r')) {
            output.WriteByte((byte)'\n');
        }
    }

    private static StreamWriter CreateWriter(Stream output) =>
        new(output, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };

    private sealed class SerializedObject {
        public SerializedObject(int objectNumber, int generation, byte[] bytes) {
            ObjectNumber = objectNumber;
            Generation = generation;
            Bytes = bytes;
        }

        public int ObjectNumber { get; }

        public int Generation { get; }

        public byte[] Bytes { get; }
    }
}
