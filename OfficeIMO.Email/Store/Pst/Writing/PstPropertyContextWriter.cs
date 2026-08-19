using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstWriterObjectReference {
    internal PstWriterObjectReference(uint nid, long size) {
        Nid = nid;
        Size = size;
    }

    internal uint Nid { get; }
    internal long Size { get; }
}

internal sealed class PstWriterValueReference {
    internal PstWriterValueReference(uint nid, ulong dataBid, ulong subnodeBid = 0) {
        Nid = nid;
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
    }

    internal uint Nid { get; }
    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
}

internal readonly struct PstWriterContextResult {
    internal PstWriterContextResult(ulong dataBid, ulong subnodeBid,
        long serializedDataLength = 0) {
        DataBid = dataBid;
        SubnodeBid = subnodeBid;
        SerializedDataLength = serializedDataLength;
    }

    internal ulong DataBid { get; }
    internal ulong SubnodeBid { get; }
    internal long SerializedDataLength { get; }
}

internal static class PstPropertyContextWriter {
    private const int MaximumHeapPayload = 8176;
    private const int PreferredHeapValueLimit = 3580;

    internal static PstWriterContextResult Write(PstWriterFile file,
        IEnumerable<MapiProperty> properties, int codePage,
        IEnumerable<PstWriterSubnode>? additionalSubnodes,
        IReadOnlyDictionary<ushort, PstWriterValueReference>? valueReferences,
        IReadOnlyDictionary<ushort, PstWriterObjectReference>? objectReferences,
        Action<EmailStoreDiagnostic>? reportDiagnostic, string location) =>
        WriteCore(file, properties, codePage, additionalSubnodes, valueReferences,
            objectReferences, reportDiagnostic, location, null);

    internal static PstWriterContextResult WriteWithCalculatedSize(PstWriterFile file,
        IEnumerable<MapiProperty> properties, int codePage,
        IEnumerable<PstWriterSubnode>? additionalSubnodes,
        IReadOnlyDictionary<ushort, PstWriterValueReference>? valueReferences,
        IReadOnlyDictionary<ushort, PstWriterObjectReference>? objectReferences,
        Action<EmailStoreDiagnostic>? reportDiagnostic, string location,
        MapiPropertyKey<int> sizeProperty) =>
        WriteCore(file, properties, codePage, additionalSubnodes, valueReferences,
            objectReferences, reportDiagnostic, location,
            sizeProperty?.GetStandardPropertyId() ?? throw new ArgumentNullException(nameof(sizeProperty)));

    private static PstWriterContextResult WriteCore(PstWriterFile file,
        IEnumerable<MapiProperty> properties, int codePage,
        IEnumerable<PstWriterSubnode>? additionalSubnodes,
        IReadOnlyDictionary<ushort, PstWriterValueReference>? valueReferences,
        IReadOnlyDictionary<ushort, PstWriterObjectReference>? objectReferences,
        Action<EmailStoreDiagnostic>? reportDiagnostic, string location,
        ushort? calculatedSizePropertyId) {
        if (file == null) throw new ArgumentNullException(nameof(file));
        if (properties == null) throw new ArgumentNullException(nameof(properties));

        var candidates = new List<EncodedProperty>();
        foreach (MapiProperty property in properties
            .Where(item => item != null)
            .GroupBy(item => item.PropertyId)
            .Select(group => group.Last())
            .OrderBy(item => item.PropertyId)) {
            try {
                if (valueReferences != null && valueReferences.TryGetValue(
                    property.PropertyId, out PstWriterValueReference? valueReference)) {
                    candidates.Add(new EncodedProperty(property.PropertyId,
                        property.PropertyType, valueReference.Nid, isExternalReference: true));
                } else if (property.PropertyType == MapiPropertyType.Object) {
                    if (objectReferences != null && objectReferences.TryGetValue(
                        property.PropertyId, out PstWriterObjectReference? reference)) {
                        var descriptor = new byte[8];
                        PstBinary.WriteUInt32(descriptor, 0, reference.Nid);
                        PstBinary.WriteUInt32(descriptor, 4, checked((uint)Math.Min(reference.Size, uint.MaxValue)));
                        candidates.Add(new EncodedProperty(property.PropertyId,
                            property.PropertyType, descriptor, isObjectDescriptor: true));
                    } else {
                        Report(reportDiagnostic, "EMAIL_STORE_PST_WRITE_OBJECT_OMITTED",
                            "An object property without a corresponding embedded subnode was omitted.",
                            EmailStoreDiagnosticSeverity.Warning, location);
                    }
                } else if (PstPropertyValueWriter.IsInline(property.PropertyType)) {
                    candidates.Add(new EncodedProperty(property.PropertyId,
                        property.PropertyType, PstPropertyValueWriter.EncodeInline(property)));
                } else {
                    candidates.Add(new EncodedProperty(property.PropertyId,
                        property.PropertyType,
                        PstPropertyValueWriter.EncodeVariable(property, codePage)));
                }
            } catch (Exception exception) when (exception is ArgumentException ||
                exception is InvalidCastException || exception is FormatException ||
                exception is OverflowException || exception is NotSupportedException) {
                Report(reportDiagnostic, "EMAIL_STORE_PST_WRITE_PROPERTY_OMITTED",
                    string.Concat("Property 0x", property.PropertyTag.ToString("X8", CultureInfo.InvariantCulture),
                        " was omitted: ", exception.Message),
                    EmailStoreDiagnosticSeverity.Warning, location);
            }
        }

        foreach (EncodedProperty candidate in candidates.Where(item =>
            item.Bytes != null && !item.IsObjectDescriptor && item.Bytes.Length > PreferredHeapValueLimit)) {
            candidate.UseSubnode = true;
        }
        var subnodes = additionalSubnodes == null
            ? new List<PstWriterSubnode>()
            : new List<PstWriterSubnode>(additionalSubnodes);
        if (valueReferences != null) {
            foreach (PstWriterValueReference reference in valueReferences.Values) {
                subnodes.Add(new PstWriterSubnode(reference.Nid, reference.DataBid, reference.SubnodeBid));
            }
        }
        uint localIndex = NextLocalIndex(subnodes);
        foreach (EncodedProperty candidate in candidates.Where(item => item.UseSubnode)) {
            candidate.SubnodeNid = checked((localIndex++ << 5) | 0x1FU);
        }
        IReadOnlyList<byte[]> heapBlocks = BuildHeap(candidates);
        long serializedDataLength = heapBlocks.Sum(item => (long)item.Length);
        foreach (EncodedProperty candidate in candidates.Where(item => item.UseSubnode)) {
            serializedDataLength = checked(serializedDataLength + candidate.Bytes!.LongLength);
        }
        foreach (PstWriterSubnode subnode in subnodes) {
            serializedDataLength = checked(serializedDataLength +
                file.GetContextDataLength(subnode.DataBid, subnode.SubnodeBid));
        }
        if (calculatedSizePropertyId.HasValue) {
            EncodedProperty size = candidates.Single(item =>
                item.PropertyId == calculatedSizePropertyId.Value &&
                item.PropertyType == MapiPropertyType.Integer32 && item.IsInline);
            size.InlineValue = checked((uint)Math.Min(serializedDataLength, int.MaxValue));
            heapBlocks = BuildHeap(candidates);
        }
        foreach (EncodedProperty candidate in candidates.Where(item => item.UseSubnode)) {
            subnodes.Add(new PstWriterSubnode(candidate.SubnodeNid,
                file.WriteDataTree(candidate.Bytes!)));
        }
        ulong dataBid = file.WriteDataTreeBlocks(heapBlocks);
        ulong subnodeBid = PstWriterSubnodeTree.Write(file, subnodes);
        return new PstWriterContextResult(dataBid, subnodeBid, serializedDataLength);
    }

    private static IReadOnlyList<byte[]> BuildHeap(IReadOnlyList<EncodedProperty> candidates) {
        var heap = new PstWriterHeap(0xBC);
        var bthHeader = new byte[8];
        uint headerHid = heap.Add(bthHeader);
        var records = new byte[candidates.Count * 8];
        for (int index = 0; index < candidates.Count; index++) {
            EncodedProperty candidate = candidates[index];
            uint rawValue;
            if (candidate.IsExternalReference || candidate.IsInline) rawValue = candidate.InlineValue;
            else if (candidate.Bytes == null || candidate.Bytes.Length == 0) rawValue = 0;
            else if (candidate.UseSubnode) rawValue = candidate.SubnodeNid;
            else rawValue = heap.Add(candidate.Bytes);
            int recordOffset = index * 8;
            PstBinary.WriteUInt16(records, recordOffset, candidate.PropertyId);
            PstBinary.WriteUInt16(records, recordOffset + 2, (ushort)candidate.PropertyType);
            PstBinary.WriteUInt32(records, recordOffset + 4, rawValue);
        }
        PstWriterBth.Complete(heap, bthHeader, 2, 6, records);
        return heap.Build(headerHid);
    }

    private static uint NextLocalIndex(IEnumerable<PstWriterSubnode> subnodes) {
        uint maximum = 0x10;
        foreach (PstWriterSubnode subnode in subnodes) maximum = Math.Max(maximum, subnode.Nid >> 5);
        return checked(maximum + 1);
    }

    private static void Report(Action<EmailStoreDiagnostic>? callback, string code,
        string message, EmailStoreDiagnosticSeverity severity, string location) =>
        callback?.Invoke(new EmailStoreDiagnostic(code, message, severity, location));

    private sealed class EncodedProperty {
        internal EncodedProperty(ushort propertyId, MapiPropertyType propertyType, uint inlineValue,
            bool isExternalReference = false) {
            PropertyId = propertyId;
            PropertyType = propertyType;
            InlineValue = inlineValue;
            IsInline = !isExternalReference;
            IsExternalReference = isExternalReference;
        }

        internal EncodedProperty(ushort propertyId, MapiPropertyType propertyType, byte[] bytes,
            bool isObjectDescriptor = false) {
            PropertyId = propertyId;
            PropertyType = propertyType;
            Bytes = bytes;
            IsObjectDescriptor = isObjectDescriptor;
        }

        internal ushort PropertyId { get; }
        internal MapiPropertyType PropertyType { get; }
        internal uint InlineValue { get; set; }
        internal bool IsInline { get; }
        internal bool IsExternalReference { get; }
        internal byte[]? Bytes { get; }
        internal bool IsObjectDescriptor { get; }
        internal bool UseSubnode { get; set; }
        internal uint SubnodeNid { get; set; }
    }
}
