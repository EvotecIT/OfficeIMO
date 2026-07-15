using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Assembles a classic-cross-reference PDF using the two-section layout and hint tables defined by PDF 1.7 Appendix F.</summary>
internal static class PdfLinearizationFileAssembler {
    private const int FixedIntegerWidth = 20;

    internal static byte[] Assemble(
        Dictionary<int, PdfIndirectObject> objects,
        int catalogObjectNumber,
        PdfMetadata metadata,
        byte[] sourcePdf) {
        PdfLinearizationPlan plan = PdfLinearizationPlanner.Create(objects, catalogObjectNumber);
        Numbering numbering = CreateNumbering(plan);
        SerializedObjects serialized = Serialize(objects, metadata, plan, numbering);

        byte[] header = PdfEncoding.Latin1GetBytes(
            "%PDF-" + PdfFileAssembler.GetHeaderVersion(PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(sourcePdf))) + "\n%\u00e2\u00e3\u00cf\u00d3\n");

        byte[] placeholderLinearization = BuildLinearizationDictionary(numbering.LinearizationObjectId, 0, 0, 0, numbering.FirstPageObjectId, 0, plan.PageObjectNumbers.Count, 0);
        byte[] placeholderFirstXref = BuildFirstPageXref(
            numbering.FirstGroupStart,
            numbering.FirstGroupCount,
            new Dictionary<int, long>(),
            numbering.TotalObjectCount + 1,
            0,
            numbering.CatalogObjectId,
            numbering.InfoObjectId);

        long linearizationOffset = header.LongLength;
        long firstXrefOffset = linearizationOffset + placeholderLinearization.LongLength;
        long catalogOffset = firstXrefOffset + placeholderFirstXref.LongLength;
        long hintOffset = catalogOffset + serialized.Catalog.LongLength;

        // Hint-table offsets at or after /H are stored as if the hint stream were absent;
        // readers add its declared length when resolving them (PDF 1.7, F.3).
        long rawFirstPageOffset = hintOffset;
        PageHintData pageHints = BuildPageHints(plan, numbering, serialized, rawFirstPageOffset);
        byte[] hintData = PdfObjectBytes.Concat(pageHints.PageOffsetTable, pageHints.SharedObjectTable);
        string hintDictionary = "<< /Length " + hintData.Length.ToString(CultureInfo.InvariantCulture) +
            " /S " + pageHints.PageOffsetTable.Length.ToString(CultureInfo.InvariantCulture) + " >>";
        byte[] hintObject = PdfObjectBytes.WrapStreamObject(numbering.HintObjectId, hintDictionary, hintData);
        long hintLength = hintObject.LongLength;

        long firstPageOffset = rawFirstPageOffset + hintLength;
        var objectOffsets = new Dictionary<int, long> {
            [numbering.LinearizationObjectId] = linearizationOffset,
            [numbering.CatalogObjectId] = catalogOffset,
            [numbering.HintObjectId] = hintOffset
        };

        long position = firstPageOffset;
        RecordOffsets(serialized.PageGroups[0], objectOffsets, ref position);
        long firstPageEnd = position;
        for (int pageIndex = 1; pageIndex < serialized.PageGroups.Count; pageIndex++) {
            RecordOffsets(serialized.PageGroups[pageIndex], objectOffsets, ref position);
        }

        RecordOffsets(serialized.SharedObjects, objectOffsets, ref position);
        RecordOffsets(serialized.RemainingObjects, objectOffsets, ref position);
        objectOffsets[numbering.InfoObjectId] = position;
        position += serialized.Info.LongLength;
        long mainXrefOffset = position;
        byte[] mainXref = BuildMainXref(numbering.FirstGroupStart, objectOffsets, firstXrefOffset);
        long fileLength = mainXrefOffset + mainXref.LongLength;
        long mainXrefFirstEntryOffset = mainXrefOffset + (
            "xref\n0 " + numbering.FirstGroupStart.ToString(CultureInfo.InvariantCulture) + "\n").Length;

        byte[] linearization = BuildLinearizationDictionary(
            numbering.LinearizationObjectId,
            fileLength,
            hintOffset,
            hintLength,
            numbering.FirstPageObjectId,
            firstPageEnd,
            plan.PageObjectNumbers.Count,
            mainXrefFirstEntryOffset);
        byte[] firstXref = BuildFirstPageXref(
            numbering.FirstGroupStart,
            numbering.FirstGroupCount,
            objectOffsets,
            numbering.TotalObjectCount + 1,
            mainXrefOffset,
            numbering.CatalogObjectId,
            numbering.InfoObjectId);
        if (linearization.Length != placeholderLinearization.Length || firstXref.Length != placeholderFirstXref.Length) {
            throw new InvalidOperationException("Linearized PDF fixed-width planning changed between assembly passes.");
        }

        using var output = new MemoryStream(checked((int)fileLength));
        Write(output, header);
        Write(output, linearization);
        Write(output, firstXref);
        Write(output, serialized.Catalog);
        Write(output, hintObject);
        WriteGroup(output, serialized.PageGroups[0]);
        for (int pageIndex = 1; pageIndex < serialized.PageGroups.Count; pageIndex++) WriteGroup(output, serialized.PageGroups[pageIndex]);
        WriteGroup(output, serialized.SharedObjects);
        WriteGroup(output, serialized.RemainingObjects);
        Write(output, serialized.Info);
        Write(output, mainXref);
        if (output.Position != fileLength) throw new InvalidOperationException("Linearized PDF output length did not match its /L parameter.");
        return output.ToArray();
    }

    private static Numbering CreateNumbering(PdfLinearizationPlan plan) {
        var numberMap = new Dictionary<int, int>();
        int next = 1;
        for (int pageIndex = 1; pageIndex < plan.PageGroups.Count; pageIndex++) {
            foreach (int sourceId in plan.PageGroups[pageIndex]) numberMap[sourceId] = next++;
        }
        foreach (int sourceId in plan.SharedObjects) numberMap[sourceId] = next++;
        foreach (int sourceId in plan.RemainingObjects) numberMap[sourceId] = next++;
        int infoObjectId = next++;
        int firstGroupStart = next;
        int linearizationObjectId = next++;
        int catalogObjectId = next++;
        numberMap[plan.CatalogObjectNumber] = catalogObjectId;
        foreach (int sourceId in plan.PageGroups[0]) numberMap[sourceId] = next++;
        int firstPageObjectId = numberMap[plan.PageObjectNumbers[0]];
        int hintObjectId = next++;
        return new Numbering(numberMap, infoObjectId, firstGroupStart, linearizationObjectId, catalogObjectId, firstPageObjectId, hintObjectId, next - 1);
    }

    private static SerializedObjects Serialize(
        Dictionary<int, PdfIndirectObject> objects,
        PdfMetadata metadata,
        PdfLinearizationPlan plan,
        Numbering numbering) {
        var context = new PdfPageExtractor.SerializationContext(
            numbering.NumberMap,
            pagesObjectId: 0,
            new Dictionary<int, Dictionary<string, PdfObject>>(),
            objects);

        byte[] SerializeObject(int sourceId) => PdfObjectBytes.WrapIndirectObject(
            numbering.NumberMap[sourceId],
            PdfPageExtractor.SerializeObject(objects[sourceId].Value, context));

        var pageGroups = new List<IReadOnlyList<SerializedObject>>(plan.PageGroups.Count);
        foreach (IReadOnlyList<int> sourceGroup in plan.PageGroups) {
            pageGroups.Add(sourceGroup.Select(sourceId => new SerializedObject(numbering.NumberMap[sourceId], sourceId, SerializeObject(sourceId))).ToList().AsReadOnly());
        }

        var shared = plan.SharedObjects.Select(sourceId => new SerializedObject(numbering.NumberMap[sourceId], sourceId, SerializeObject(sourceId))).ToList().AsReadOnly();
        var remaining = plan.RemainingObjects.Select(sourceId => new SerializedObject(numbering.NumberMap[sourceId], sourceId, SerializeObject(sourceId))).ToList().AsReadOnly();
        byte[] catalog = SerializeObject(plan.CatalogObjectNumber);
        byte[] info = PdfObjectBytes.WrapIndirectObject(numbering.InfoObjectId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(metadata)));
        return new SerializedObjects(catalog, pageGroups.AsReadOnly(), shared, remaining, info);
    }

    private static PageHintData BuildPageHints(
        PdfLinearizationPlan plan,
        Numbering numbering,
        SerializedObjects serialized,
        long rawFirstPageOffset) {
        int pageCount = serialized.PageGroups.Count;
        var objectCounts = new uint[pageCount];
        var pageLengths = new uint[pageCount];
        var sharedIdentifiers = new List<uint>[pageCount];
        var sharedIdentifierBySourceId = new Dictionary<int, uint>();
        var sharedGroupLengths = new List<uint>();

        uint sharedIndex = 0;
        foreach (SerializedObject item in serialized.PageGroups[0]) {
            sharedIdentifierBySourceId[item.SourceId] = sharedIndex++;
            sharedGroupLengths.Add(checked((uint)item.Bytes.Length));
        }
        foreach (SerializedObject item in serialized.SharedObjects) {
            sharedIdentifierBySourceId[item.SourceId] = sharedIndex++;
            sharedGroupLengths.Add(checked((uint)item.Bytes.Length));
        }

        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++) {
            objectCounts[pageIndex] = checked((uint)serialized.PageGroups[pageIndex].Count);
            pageLengths[pageIndex] = checked((uint)serialized.PageGroups[pageIndex].Sum(static item => item.Bytes.Length));
            var identifiers = new List<uint>();
            if (pageIndex > 0) {
                foreach (SerializedObject item in serialized.PageGroups[0]) {
                    if (plan.ReachableByPage[pageIndex].Contains(item.SourceId)) identifiers.Add(sharedIdentifierBySourceId[item.SourceId]);
                }
                foreach (SerializedObject item in serialized.SharedObjects) {
                    if (plan.ReachableByPage[pageIndex].Contains(item.SourceId)) identifiers.Add(sharedIdentifierBySourceId[item.SourceId]);
                }
            }
            sharedIdentifiers[pageIndex] = identifiers;
        }

        byte[] pageOffsetTable = BuildPageOffsetHintTable(objectCounts, pageLengths, sharedIdentifiers, checked((uint)rawFirstPageOffset));
        long rawSharedOffset = rawFirstPageOffset + pageLengths.Aggregate(0L, static (sum, length) => sum + length);
        byte[] sharedObjectTable = BuildSharedObjectHintTable(
            sharedGroupLengths,
            checked((uint)serialized.PageGroups[0].Count),
            serialized.SharedObjects.Count > 0 ? checked((uint)serialized.SharedObjects[0].ObjectId) : 0U,
            serialized.SharedObjects.Count > 0 ? checked((uint)rawSharedOffset) : 0U);
        return new PageHintData(pageOffsetTable, sharedObjectTable);
    }

    private static byte[] BuildPageOffsetHintTable(
        uint[] objectCounts,
        uint[] pageLengths,
        IReadOnlyList<uint>[] sharedIdentifiers,
        uint firstPageOffset) {
        uint minObjects = objectCounts.Min();
        uint minPageLength = pageLengths.Min();
        int objectDeltaBits = BitsRequired(objectCounts.Max() - minObjects);
        int pageLengthDeltaBits = BitsRequired(pageLengths.Max() - minPageLength);
        int sharedCountBits = BitsRequired(checked((uint)sharedIdentifiers.Max(static identifiers => identifiers.Count)));
        uint maxSharedIdentifier = sharedIdentifiers.SelectMany(static identifiers => identifiers).DefaultIfEmpty(0U).Max();
        int sharedIdentifierBits = sharedIdentifiers.Any(static identifiers => identifiers.Count > 0) ? BitsRequired(maxSharedIdentifier) : 0;

        var writer = new PdfLinearizationBitWriter();
        writer.Write(minObjects, 32);
        writer.Write(firstPageOffset, 32);
        writer.Write((uint)objectDeltaBits, 16);
        writer.Write(minPageLength, 32);
        writer.Write((uint)pageLengthDeltaBits, 16);
        writer.Write(0U, 32);
        writer.Write(0U, 16);
        writer.Write(minPageLength, 32);
        writer.Write((uint)pageLengthDeltaBits, 16);
        writer.Write((uint)sharedCountBits, 16);
        writer.Write((uint)sharedIdentifierBits, 16);
        writer.Write(0U, 16);
        writer.Write(4U, 16);

        for (int index = 0; index < objectCounts.Length; index++) writer.Write(objectCounts[index] - minObjects, objectDeltaBits);
        writer.AlignToByte();
        for (int index = 0; index < pageLengths.Length; index++) writer.Write(pageLengths[index] - minPageLength, pageLengthDeltaBits);
        writer.AlignToByte();
        for (int index = 0; index < sharedIdentifiers.Length; index++) writer.Write(checked((uint)sharedIdentifiers[index].Count), sharedCountBits);
        writer.AlignToByte();
        for (int index = 1; index < sharedIdentifiers.Length; index++) {
            foreach (uint identifier in sharedIdentifiers[index]) writer.Write(identifier, sharedIdentifierBits);
        }
        writer.AlignToByte();
        for (int index = 0; index < objectCounts.Length; index++) writer.Write(0U, 0);
        writer.AlignToByte();
        for (int index = 0; index < objectCounts.Length; index++) writer.Write(0U, 0);
        writer.AlignToByte();
        for (int index = 0; index < pageLengths.Length; index++) writer.Write(pageLengths[index] - minPageLength, pageLengthDeltaBits);
        writer.AlignToByte();
        return writer.ToArray();
    }

    private static byte[] BuildSharedObjectHintTable(
        List<uint> groupLengths,
        uint firstPageEntryCount,
        uint firstSharedObjectNumber,
        uint firstSharedOffset) {
        if (groupLengths.Count == 0) throw new InvalidOperationException("Linearized PDF first-page group cannot be empty.");
        uint minLength = groupLengths.Min();
        int lengthDeltaBits = BitsRequired(groupLengths.Max() - minLength);
        var writer = new PdfLinearizationBitWriter();
        writer.Write(firstSharedObjectNumber, 32);
        writer.Write(firstSharedOffset, 32);
        writer.Write(firstPageEntryCount, 32);
        writer.Write(checked((uint)groupLengths.Count), 32);
        writer.Write(0U, 16);
        writer.Write(minLength, 32);
        writer.Write((uint)lengthDeltaBits, 16);
        foreach (uint length in groupLengths) writer.Write(length - minLength, lengthDeltaBits);
        writer.AlignToByte();
        foreach (uint _ in groupLengths) writer.Write(0U, 1);
        writer.AlignToByte();
        writer.AlignToByte();
        writer.AlignToByte();
        return writer.ToArray();
    }

    private static byte[] BuildLinearizationDictionary(
        int objectId,
        long fileLength,
        long hintOffset,
        long hintLength,
        int firstPageObjectId,
        long firstPageEnd,
        int pageCount,
        long mainXrefFirstEntryOffset) {
        string body = "<< /Linearized 1 /L " + Fixed(fileLength) +
            " /H [ " + Fixed(hintOffset) + " " + Fixed(hintLength) + " ] /O " + firstPageObjectId.ToString(CultureInfo.InvariantCulture) +
            " /E " + Fixed(firstPageEnd) + " /N " + pageCount.ToString(CultureInfo.InvariantCulture) +
            " /T " + Fixed(mainXrefFirstEntryOffset) + " >>\n";
        byte[] result = PdfObjectBytes.WrapIndirectObject(objectId, body);
        if (result.Length > 1024) throw new InvalidOperationException("Linearization parameter dictionary exceeds the first 1024 bytes of the PDF file.");
        return result;
    }

    private static byte[] BuildFirstPageXref(
        int firstGroupStart,
        int firstGroupCount,
        Dictionary<int, long> offsets,
        int totalSize,
        long mainXrefOffset,
        int catalogObjectId,
        int infoObjectId) {
        var builder = new StringBuilder();
        builder.Append("xref\n").Append(firstGroupStart.ToString(CultureInfo.InvariantCulture)).Append(' ')
            .Append(firstGroupCount.ToString(CultureInfo.InvariantCulture)).Append('\n');
        for (int objectId = firstGroupStart; objectId < firstGroupStart + firstGroupCount; objectId++) {
            offsets.TryGetValue(objectId, out long offset);
            AppendXrefEntry(builder, offset, inUse: true);
        }

        builder.Append("trailer\n<< /Size ").Append(totalSize.ToString(CultureInfo.InvariantCulture))
            .Append(" /Prev ").Append(Fixed(mainXrefOffset))
            .Append(" /Root ").Append(PdfSyntaxEscaper.IndirectReference(catalogObjectId))
            .Append(" /Info ").Append(PdfSyntaxEscaper.IndirectReference(infoObjectId))
            .Append(" >>\nstartxref\n0\n%%EOF\n");
        return PdfEncoding.Latin1GetBytes(builder.ToString());
    }

    private static byte[] BuildMainXref(int firstGroupStart, Dictionary<int, long> offsets, long firstXrefOffset) {
        var builder = new StringBuilder();
        builder.Append("xref\n0 ").Append(firstGroupStart.ToString(CultureInfo.InvariantCulture)).Append('\n');
        AppendXrefEntry(builder, 0, inUse: false);
        for (int objectId = 1; objectId < firstGroupStart; objectId++) {
            if (!offsets.TryGetValue(objectId, out long offset)) throw new InvalidOperationException("Linearized PDF main xref is missing object " + objectId.ToString(CultureInfo.InvariantCulture) + ".");
            AppendXrefEntry(builder, offset, inUse: true);
        }
        builder.Append("trailer\n<< /Size ").Append(firstGroupStart.ToString(CultureInfo.InvariantCulture))
            .Append(" >>\nstartxref\n").Append(firstXrefOffset.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return PdfEncoding.Latin1GetBytes(builder.ToString());
    }

    private static void AppendXrefEntry(StringBuilder builder, long offset, bool inUse) {
        if (offset < 0 || offset > 9999999999L) throw new NotSupportedException("Classic PDF cross-reference offsets cannot exceed ten decimal digits.");
        builder.Append(offset.ToString("0000000000", CultureInfo.InvariantCulture))
            .Append(inUse ? " 00000 n \n" : " 65535 f \n");
    }

    private static void RecordOffsets(IReadOnlyList<SerializedObject> objects, Dictionary<int, long> offsets, ref long position) {
        foreach (SerializedObject item in objects) {
            offsets[item.ObjectId] = position;
            position += item.Bytes.LongLength;
        }
    }

    private static void WriteGroup(Stream destination, IReadOnlyList<SerializedObject> objects) {
        foreach (SerializedObject item in objects) Write(destination, item.Bytes);
    }

    private static void Write(Stream destination, byte[] bytes) => destination.Write(bytes, 0, bytes.Length);

    #pragma warning disable CA1512 // ThrowIfNegative is unavailable on netstandard2.0 and net472.
    private static string Fixed(long value) {
        if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
        string text = value.ToString(CultureInfo.InvariantCulture);
        if (text.Length > FixedIntegerWidth) throw new NotSupportedException("Linearized PDF offset exceeds the supported fixed-width field.");
        return text.PadRight(FixedIntegerWidth, ' ');
    }
    #pragma warning restore CA1512

    private static int BitsRequired(uint value) {
        int bits = 0;
        while (value > 0U) {
            bits++;
            value >>= 1;
        }
        return bits;
    }

    private sealed class Numbering {
        internal Numbering(Dictionary<int, int> numberMap, int infoObjectId, int firstGroupStart, int linearizationObjectId, int catalogObjectId, int firstPageObjectId, int hintObjectId, int totalObjectCount) {
            NumberMap = numberMap;
            InfoObjectId = infoObjectId;
            FirstGroupStart = firstGroupStart;
            LinearizationObjectId = linearizationObjectId;
            CatalogObjectId = catalogObjectId;
            FirstPageObjectId = firstPageObjectId;
            HintObjectId = hintObjectId;
            TotalObjectCount = totalObjectCount;
        }
        internal Dictionary<int, int> NumberMap { get; }
        internal int InfoObjectId { get; }
        internal int FirstGroupStart { get; }
        internal int FirstGroupCount => TotalObjectCount - FirstGroupStart + 1;
        internal int LinearizationObjectId { get; }
        internal int CatalogObjectId { get; }
        internal int FirstPageObjectId { get; }
        internal int HintObjectId { get; }
        internal int TotalObjectCount { get; }
    }

    private sealed class SerializedObjects {
        internal SerializedObjects(byte[] catalog, IReadOnlyList<IReadOnlyList<SerializedObject>> pageGroups, IReadOnlyList<SerializedObject> sharedObjects, IReadOnlyList<SerializedObject> remainingObjects, byte[] info) {
            Catalog = catalog; PageGroups = pageGroups; SharedObjects = sharedObjects; RemainingObjects = remainingObjects; Info = info;
        }
        internal byte[] Catalog { get; }
        internal IReadOnlyList<IReadOnlyList<SerializedObject>> PageGroups { get; }
        internal IReadOnlyList<SerializedObject> SharedObjects { get; }
        internal IReadOnlyList<SerializedObject> RemainingObjects { get; }
        internal byte[] Info { get; }
    }

    private readonly struct SerializedObject {
        internal SerializedObject(int objectId, int sourceId, byte[] bytes) { ObjectId = objectId; SourceId = sourceId; Bytes = bytes; }
        internal int ObjectId { get; }
        internal int SourceId { get; }
        internal byte[] Bytes { get; }
    }

    private readonly struct PageHintData {
        internal PageHintData(byte[] pageOffsetTable, byte[] sharedObjectTable) { PageOffsetTable = pageOffsetTable; SharedObjectTable = sharedObjectTable; }
        internal byte[] PageOffsetTable { get; }
        internal byte[] SharedObjectTable { get; }
    }
}
