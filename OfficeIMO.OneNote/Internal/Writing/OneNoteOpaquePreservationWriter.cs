namespace OfficeIMO.OneNote;

/// <summary>Merges retained public opaque objects into regenerated TOC object spaces.</summary>
internal static class OneNoteOpaquePreservationWriter {
    internal static void MergeSpace(
        OneNoteWriteObjectSpace generated,
        IEnumerable<OneNoteOpaqueObject> preservedObjects,
        long maxPayloadBytes) {
        if (generated == null) throw new ArgumentNullException(nameof(generated));
        if (preservedObjects == null) throw new ArgumentNullException(nameof(preservedObjects));
        if (maxPayloadBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxPayloadBytes));

        var generatedById = generated.Objects
            .GroupBy(item => item.Id)
            .ToDictionary(group => group.Key, group => group.Last());
        var merged = new List<OneNoteWriteObject>();
        var emitted = new HashSet<OneNoteExtendedGuid>();

        foreach (OneNoteOpaqueObject source in preservedObjects.OrderBy(item => item.Ordinal)) {
            OneNoteExtendedGuid id = source.Id ?? new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
            source.Id = id;
            if (!emitted.Add(id)) continue;
            OneNoteWriteObject retained = ConvertObject(source, id, maxPayloadBytes);
            merged.Add(generatedById.TryGetValue(id, out OneNoteWriteObject? typed)
                ? MergeObject(typed, retained)
                : retained);
        }

        foreach (OneNoteWriteObject typed in generated.Objects) {
            if (emitted.Add(typed.Id)) merged.Add(typed);
        }

        generated.Objects.Clear();
        foreach (OneNoteWriteObject item in merged) generated.Objects.Add(item);
    }

    private static OneNoteWriteObject MergeObject(OneNoteWriteObject generated, OneNoteWriteObject retained) {
        var generatedByKey = generated.Properties
            .GroupBy(PropertyKey)
            .ToDictionary(group => group.Key, group => group.Last());
        var emitted = new HashSet<uint>();
        var properties = new List<OneNoteWriteProperty>();

        foreach (OneNoteWriteProperty original in retained.Properties) {
            uint key = PropertyKey(original);
            if (generatedByKey.TryGetValue(key, out OneNoteWriteProperty? replacement)) {
                if (emitted.Add(key)) properties.Add(replacement);
            } else {
                properties.Add(original);
            }
        }
        foreach (OneNoteWriteProperty property in generated.Properties) {
            if (emitted.Add(PropertyKey(property))) properties.Add(property);
        }
        return new OneNoteWriteObject(generated.Id, generated.Jcid, properties);
    }

    private static OneNoteWriteObject ConvertObject(
        OneNoteOpaqueObject source,
        OneNoteExtendedGuid id,
        long maxPayloadBytes) {
        IReadOnlyList<OneNoteWriteProperty> properties = source.Properties
            .OrderBy(property => property.Ordinal)
            .Select(property => ConvertProperty(property, maxPayloadBytes))
            .ToArray();
        return new OneNoteWriteObject(id, source.Jcid, properties);
    }

    private static OneNoteWriteProperty ConvertProperty(OneNoteOpaqueProperty source, long maxPayloadBytes) {
        if (source.ValueType == OneNotePropertyValueType.PropertySet || source.ValueType == OneNotePropertyValueType.PropertySetArray) {
            throw new OneNoteFormatException(
                "ONENOTE_WRITE_OPAQUE_PROPERTY_SET",
                "An opaque TOC object contains a nested property set that cannot yet be regenerated safely.");
        }

        byte[] data = source.GetRawData();
        if (data.LongLength > maxPayloadBytes) throw new IOException("An opaque OneNote property exceeds MaxOutputBytes.");
        OneNoteWriteReferenceKind referenceKind;
        switch (source.ValueType) {
            case OneNotePropertyValueType.ObjectSpaceId:
            case OneNotePropertyValueType.ObjectSpaceIdArray:
                referenceKind = OneNoteWriteReferenceKind.ObjectSpace;
                break;
            case OneNotePropertyValueType.ContextId:
            case OneNotePropertyValueType.ContextIdArray:
                referenceKind = OneNoteWriteReferenceKind.Context;
                break;
            default:
                referenceKind = OneNoteWriteReferenceKind.Object;
                break;
        }
        return new OneNoteWriteProperty(
            source.PropertyId,
            data,
            source.ScalarValue,
            source.BooleanValue,
            source.ReferencedIds,
            referenceKind,
            preserveRawId: true);
    }

    private static uint PropertyKey(OneNoteWriteProperty property) => property.RawId & 0x03FFFFFFU;
}
