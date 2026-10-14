namespace OfficeIMO.IWork.Internal;

/// <summary>Decodes the same native padding and alignment properties for selected and region cell styles.</summary>
internal sealed class IWorkCellLayoutReader(IWorkSourceReferenceIssueCollector references) {
    internal void ReadPadding(IWorkArchiveRecord owner, IWorkWireMessage properties,
        ref IWorkCellPadding? padding, ref bool complete) {
        if (!properties.HasField(9)) return;
        IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, 9, out bool malformed);
        bool valid = !malformed && properties.FieldCount(9) == 1
            && !properties.HasUnexpectedWireKind(9, IWorkWireKind.Bytes) && declaration != null;
        var sides = new double[4];
        if (valid) {
            valid = declaration!.TotalFieldCount == Enumerable.Range(1, 4).Sum(declaration.FieldCount);
            for (int field = 1; field <= 4; field++) {
                float? value = declaration.GetFloat(field);
                if (declaration.FieldCount(field) > 1
                    || declaration.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
                    || declaration.HasField(field) && !value.HasValue
                    || value.HasValue && (float.IsNaN(value.Value) || float.IsInfinity(value.Value) || value.Value < 0))
                    valid = false;
                sides[field - 1] = value ?? 0;
            }
        }
        if (!valid) {
            references.Declarations.Record(owner, "11/9", properties.FieldCount(9),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        // The message overrides the whole property. Omitted scalar sides use protobuf zero defaults.
        padding = new IWorkCellPadding(sides[0], sides[1], sides[2], sides[3]);
    }

    internal void ReadVerticalAlignment(IWorkArchiveRecord owner, IWorkWireMessage properties,
        ref IWorkCellVerticalAlignment? vertical, ref bool complete) {
        if (!properties.HasField(8)) return;
        ulong? value = properties.GetUnsigned(8);
        if (properties.FieldCount(8) != 1 || properties.HasUnexpectedWireKind(8, IWorkWireKind.Varint)
            || !value.HasValue || value.Value > 2) {
            references.Declarations.Record(owner, "11/8", properties.FieldCount(8),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        vertical = (IWorkCellVerticalAlignment)value.Value;
    }

}

/// <summary>One resolved region layout; unknown properties remain independent.</summary>
internal sealed class IWorkCellLayout(IWorkCellPadding? padding, IWorkCellVerticalAlignment? verticalAlignment) {
    internal IWorkCellPadding? Padding { get; } = padding;
    internal IWorkCellVerticalAlignment? VerticalAlignment { get; } = verticalAlignment;
}
