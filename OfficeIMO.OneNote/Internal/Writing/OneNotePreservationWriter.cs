namespace OfficeIMO.OneNote;

/// <summary>
/// Merges the current typed projection with source objects that OfficeIMO did not understand.
/// Generated semantic properties win, while unknown properties, references, roots, objects, and
/// file-data payloads remain available for a loss-aware round trip.
/// </summary>
internal static class OneNotePreservationWriter {
    internal static void MergeSpace(
        OneNoteWriteObjectSpace generated,
        OneNoteMaterializedObjectSpace source,
        OneNoteSectionPreservationState preservation,
        long maxPayloadBytes) {
        if (generated == null) throw new ArgumentNullException(nameof(generated));
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (preservation == null) throw new ArgumentNullException(nameof(preservation));
        if (maxPayloadBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxPayloadBytes));

        // OneNote can share style/list objects between multiple semantic elements. The typed
        // projection may therefore generate the same retained identity more than once.
        var generatedById = generated.Objects
            .GroupBy(item => item.Id)
            .ToDictionary(group => group.Key, group => group.Last());
        var merged = new List<OneNoteWriteObject>();
        var emitted = new HashSet<OneNoteExtendedGuid>();

        foreach (OneNoteRevisionStoreObject sourceObject in source.Objects) {
            if (generatedById.TryGetValue(sourceObject.Id, out OneNoteWriteObject? typed)) {
                merged.Add(MergeObject(typed, sourceObject, preservation.MappedObjectIds, maxPayloadBytes));
            } else {
                merged.Add(ConvertObject(sourceObject, preservation.Materializer, maxPayloadBytes));
            }
            emitted.Add(sourceObject.Id);
        }

        foreach (OneNoteWriteObject typed in generatedById.Values) {
            if (emitted.Add(typed.Id)) merged.Add(typed);
        }

        generated.Objects.Clear();
        foreach (OneNoteWriteObject item in merged) generated.Objects.Add(item);
        foreach (KeyValuePair<uint, OneNoteExtendedGuid> root in source.Roots) {
            if (!generated.Roots.ContainsKey(root.Key)) generated.Roots[root.Key] = root.Value;
        }
    }

    private static OneNoteWriteObject MergeObject(
        OneNoteWriteObject generated,
        OneNoteRevisionStoreObject source,
        IReadOnlyCollection<OneNoteExtendedGuid> mappedObjectIds,
        long maxPayloadBytes) {
        IReadOnlyList<OneNoteWriteProperty> sourceProperties = ConvertProperties(source.PropertySet, maxPayloadBytes);
        if (sourceProperties.Count == 0) return generated;

        var generatedById = generated.Properties
            .GroupBy(PropertyKey)
            .ToDictionary(group => group.Key, group => group.Last());
        var emitted = new HashSet<uint>();
        var merged = new List<OneNoteWriteProperty>();

        foreach (OneNoteWriteProperty original in sourceProperties) {
            uint key = PropertyKey(original);
            if (generatedById.TryGetValue(key, out OneNoteWriteProperty? replacement)) {
                if (emitted.Add(key)) merged.Add(MergeReferences(replacement, original, mappedObjectIds));
            } else if (IsClearedTypedProperty(generated, key)) {
                continue;
            } else {
                OneNoteWriteProperty? retained = FilterRemovedTypedReferences(original, mappedObjectIds);
                if (retained != null) merged.Add(retained);
            }
        }

        foreach (OneNoteWriteProperty property in generated.Properties) {
            uint key = PropertyKey(property);
            if (emitted.Add(key)) merged.Add(property);
        }

        return new OneNoteWriteObject(
            generated.Id,
            generated.Jcid,
            merged,
            generated.Blob,
            generated.FileDataId,
            generated.FileExtension);
    }

    private static OneNoteWriteProperty MergeReferences(
        OneNoteWriteProperty generated,
        OneNoteWriteProperty source,
        IReadOnlyCollection<OneNoteExtendedGuid> mappedObjectIds) {
        if (generated.ReferenceKind != source.ReferenceKind || generated.References.Count == 0 && source.References.Count == 0) {
            return generated;
        }
        if (generated.ReferenceKind != OneNoteWriteReferenceKind.Object && IsTypedRelationship(PropertyKey(generated))) {
            return generated;
        }

        var references = new List<OneNoteExtendedGuid>(generated.References);
        var seen = new HashSet<OneNoteExtendedGuid>(references);
        foreach (OneNoteExtendedGuid id in source.References) {
            if (!mappedObjectIds.Contains(id) && seen.Add(id)) references.Add(id);
        }
        return CloneWithReferences(generated, references);
    }

    private static OneNoteWriteProperty? FilterRemovedTypedReferences(
        OneNoteWriteProperty source,
        IReadOnlyCollection<OneNoteExtendedGuid> mappedObjectIds) {
        if (!IsTypedRelationship(PropertyKey(source)) || source.References.Count == 0) return source;
        if (source.ReferenceKind != OneNoteWriteReferenceKind.Object) return null;
        OneNoteExtendedGuid[] references = source.References.Where(id => !mappedObjectIds.Contains(id)).ToArray();
        if (references.Length == 0) return null;
        return CloneWithReferences(source, references);
    }

    private static OneNoteWriteObject ConvertObject(
        OneNoteRevisionStoreObject source,
        OneNoteObjectSpaceMaterializer materializer,
        long maxPayloadBytes) {
        byte[]? blob = null;
        Guid? fileDataId = null;
        if (source.Jcid.IsFileData && materializer.TryResolveFileData(source, out Guid id, out OneNoteBinaryPayload? payload) && payload != null) {
            blob = payload.ToArray(maxPayloadBytes);
            fileDataId = id;
        }
        return new OneNoteWriteObject(
            source.Id,
            source.Jcid.Value,
            ConvertProperties(source.PropertySet, maxPayloadBytes),
            blob,
            fileDataId,
            source.FileExtension);
    }

    private static IReadOnlyList<OneNoteWriteProperty> ConvertProperties(OneNotePropertySet? propertySet, long maxPayloadBytes) {
        if (propertySet == null) return Array.Empty<OneNoteWriteProperty>();
        return propertySet.Properties.Select(property => ConvertProperty(property, maxPayloadBytes)).ToArray();
    }

    private static OneNoteWriteProperty ConvertProperty(OneNotePropertyValue source, long maxPayloadBytes) {
        OneNoteWriteReferenceKind referenceKind;
        switch (source.Type) {
            case OneNotePropertyType.ObjectSpaceId:
            case OneNotePropertyType.ObjectSpaceIdArray:
                referenceKind = OneNoteWriteReferenceKind.ObjectSpace;
                break;
            case OneNotePropertyType.ContextId:
            case OneNotePropertyType.ContextIdArray:
                referenceKind = OneNoteWriteReferenceKind.Context;
                break;
            default:
                referenceKind = OneNoteWriteReferenceKind.Object;
                break;
        }

        return new OneNoteWriteProperty(
            source.RawPropertyId,
            source.Data?.ToArray(maxPayloadBytes),
            source.ScalarValue,
            source.BooleanValue,
            source.ReferencedIds,
            referenceKind,
            source.ChildPropertySets.Select(child => ConvertProperties(child, maxPayloadBytes)),
            source.ChildPropertyId,
            preserveRawId: true);
    }

    private static OneNoteWriteProperty CloneWithReferences(
        OneNoteWriteProperty source,
        IReadOnlyList<OneNoteExtendedGuid> references) => new OneNoteWriteProperty(
            source.RawId,
            source.Data,
            source.Scalar,
            null,
            references,
            source.ReferenceKind,
            source.ChildPropertySets,
            source.ChildPropertyId,
            preserveRawId: true);

    private static uint PropertyKey(OneNoteWriteProperty property) => property.RawId & 0x03FFFFFFU;

    private static bool IsTypedRelationship(uint propertyKey) {
        switch (propertyKey) {
            case OneNoteSchema.ContentChildNodes & 0x03FFFFFFU:
            case OneNoteSchema.ElementChildNodes & 0x03FFFFFFU:
            case OneNoteSchema.StructureElementChildNodes & 0x03FFFFFFU:
            case OneNoteSchema.ChildGraphSpaceElementNodes & 0x03FFFFFFU:
            case OneNoteSchema.VersionHistoryGraphSpaceContextNodes & 0x03FFFFFFU:
            case OneNoteSchema.MetaDataObjectsAboveGraphSpace & 0x03FFFFFFU:
            case OneNoteSchema.AuthorOriginal & 0x03FFFFFFU:
            case OneNoteSchema.AuthorMostRecent & 0x03FFFFFFU:
            case OneNoteSchema.TextRunFormatting & 0x03FFFFFFU:
            case OneNoteSchema.ListNodes & 0x03FFFFFFU:
            case OneNoteSchema.ParagraphStyle & 0x03FFFFFFU:
            case OneNoteSchema.PictureContainer & 0x03FFFFFFU:
            case OneNoteSchema.WebPictureContainer14 & 0x03FFFFFFU:
            case OneNoteSchema.EmbeddedFileContainer & 0x03FFFFFFU:
            case OneNoteSchema.PageRecognizedTextContainer & 0x03FFFFFFU:
            case OneNoteSchema.RecognizedTextChildNodes & 0x03FFFFFFU:
                return true;
            default:
                return false;
        }
    }

    private static bool IsClearedTypedProperty(OneNoteWriteObject generated, uint propertyKey) {
        bool pageMetadata = generated.Jcid == OneNoteSchema.JcidPageMetadata || generated.Jcid == OneNoteSchema.JcidConflictPageMetadata;
        if (pageMetadata && propertyKey == (OneNoteSchema.IsDeletedGraphSpaceContent & 0x03FFFFFFU)) return true;
        if (generated.Jcid != OneNoteSchema.JcidInkContainer) return false;
        bool writesInkData = generated.Properties.Any(property => PropertyKey(property) == (OneNoteSchema.InkData & 0x03FFFFFFU));
        bool writesNestedChildren = generated.Properties.Any(property => PropertyKey(property) == (OneNoteSchema.ContentChildNodes & 0x03FFFFFFU));
        return writesNestedChildren && propertyKey == (OneNoteSchema.InkData & 0x03FFFFFFU) ||
               writesInkData && propertyKey == (OneNoteSchema.ContentChildNodes & 0x03FFFFFFU);
    }
}
