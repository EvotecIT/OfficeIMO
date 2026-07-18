namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private const uint JcidNoteTagSharedDefinition = 0x00120043;

    private const uint ActionItemType = 0x10003463;
    private const uint NoteTagShape = 0x10003464;
    private const uint NoteTagHighlightColor = 0x14003465;
    private const uint NoteTagTextColor = 0x14003466;
    private const uint NoteTagPropertyStatus = 0x14003467;
    private const uint TaskTagDueDate = 0x1400346B;
    private const uint NoteTagLabel = 0x1C003468;
    private const uint NoteTagCreated = 0x1400346E;
    private const uint NoteTagCompleted = 0x1400346F;
    private const uint ActionItemStatus = 0x10003470;
    private const uint NoteTagDefinitionOid = 0x20003488;
    private const uint NoteTagStates = 0x40003489;

    private static void ApplyTags(OneNoteElement target, OneNoteRevisionStoreObject source, OneNoteMaterializedObjectSpace space) {
        OneNotePropertyValue? states = FindProperty(source.PropertySet, NoteTagStates);
        if (states == null) return;

        foreach (OneNotePropertySet state in states.ChildPropertySets.Take(9)) {
            OneNoteTag tag = MapTagState(state, space.GetObject);
            target.Tags.Add(tag);
        }
    }

    internal static OneNoteTag MapTagState(
        OneNotePropertySet state,
        Func<OneNoteExtendedGuid, OneNoteRevisionStoreObject?> resolveObject) {
        if (state == null) throw new ArgumentNullException(nameof(state));
        if (resolveObject == null) throw new ArgumentNullException(nameof(resolveObject));
        uint status = ReadUInt32(state, ActionItemStatus) ?? 0;
        uint? actionItemType = ReadUInt32(state, ActionItemType);
        uint? shape = ReadUInt32(state, NoteTagShape);
        OneNoteExtendedGuid? definitionId = GetReferences(state, NoteTagDefinitionOid).FirstOrDefault();
        OneNoteRevisionStoreObject? definition = definitionId == null ? null : resolveObject(definitionId);
        bool isTask = (status & 0x04U) != 0 || IsTaskActionItemType(actionItemType);

        if (!isTask && definition?.Jcid.Value == JcidNoteTagSharedDefinition) {
            actionItemType = ReadUInt32(definition, ActionItemType);
            shape = ReadUInt32(definition, NoteTagShape);
            isTask = IsTaskActionItemType(actionItemType);
        }

        return new OneNoteTag {
            DefinitionId = definitionId,
            ActionItemType = actionItemType,
            Label = isTask ? null : ReadString(definition, NoteTagLabel),
            Shape = shape,
            IsTask = isTask,
            IsCheckable = shape.HasValue && IsCheckableTagShape(shape.Value),
            IsCompleted = (status & 0x01U) != 0,
            IsDisabled = (status & 0x02U) != 0,
            IsUnsynchronized = (status & 0x08U) != 0,
            IsRemoved = (status & 0x10U) != 0,
            DueUtc = isTask ? ReadTime32(state, TaskTagDueDate) : null,
            CreatedUtc = ReadTime32(state, NoteTagCreated),
            CompletedUtc = ReadTime32(state, NoteTagCompleted),
            TextColorArgb = isTask ? null : ReadUInt32(definition, NoteTagTextColor),
            HighlightColorArgb = isTask ? null : ReadUInt32(definition, NoteTagHighlightColor)
        };
    }

    internal static bool IsCheckableTagShape(uint shape) {
        if (shape >= 1 && shape <= 12) return true;
        if (shape == 28 || shape == 30 || shape == 32) return true;
        if (shape == 48 || shape == 50 || shape == 52) return true;
        if (shape == 69 || shape == 71 || shape == 73) return true;
        return shape >= 89 && shape <= 99;
    }

    private static bool IsTaskActionItemType(uint? actionItemType) =>
        actionItemType.HasValue && actionItemType.Value >= 100 && actionItemType.Value <= 105;

    private static IReadOnlyList<OneNoteExtendedGuid> GetReferences(OneNotePropertySet? set, uint propertyId) {
        return FindProperty(set, propertyId)?.ReferencedIds ?? Array.Empty<OneNoteExtendedGuid>();
    }

    private static uint? ReadUInt32(OneNotePropertySet? set, uint propertyId) {
        ulong? value = FindProperty(set, propertyId)?.ScalarValue;
        return value.HasValue ? (uint)value.Value : null;
    }

    private static DateTime? ReadTime32(OneNotePropertySet? set, uint propertyId) {
        uint? value = ReadUInt32(set, propertyId);
        return value.HasValue && value.Value != 0
            ? new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(value.Value)
            : (DateTime?)null;
    }
}
