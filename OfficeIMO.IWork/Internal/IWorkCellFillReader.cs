namespace OfficeIMO.IWork.Internal;

/// <summary>Decodes the bounded FillArchive subset shared by cell, slide and text styles.</summary>
internal static class IWorkCellFillReader {
    internal static void Read(IWorkArchiveRecord owner, IWorkWireMessage properties,
        IWorkSourceReferenceIssueCollector references, ref IWorkCellFill? fill, ref bool complete, int field = 1) {
        if (!properties.HasField(field)) return;
        IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, field, out bool malformed);
        bool valid = !malformed && properties.FieldCount(field) == 1
            && !properties.HasUnexpectedWireKind(field, IWorkWireKind.Bytes) && declaration != null;
        IWorkColor? color = null;
        if (valid && declaration!.TotalFieldCount > 0) {
            valid = declaration.TotalFieldCount == 1 && declaration.FieldCount(1) == 1
                && IWorkColorReader.TryRead(declaration, 1, out color, ref valid)
                && color is { Alpha: byte.MaxValue } && IsSupportedColor(declaration);
        }
        if (!valid) {
            references.Declarations.Record(owner, "11/" + field.ToString(System.Globalization.CultureInfo.InvariantCulture), properties.FieldCount(field),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        // An empty FillArchive clears the parent rather than inheriting its color.
        fill = new IWorkCellFill(color);
    }

    private static bool IsSupportedColor(IWorkWireMessage fill) {
        IWorkWireMessage color = IWorkObjectIndex.TryGetMessage(fill, 1)!;
        int[] fields = { 1, 3, 4, 5, 6, 11, 12, 13 };
        if (color.TotalFieldCount != fields.Sum(color.FieldCount)) return false;
        foreach (int field in new[] { 1, 12 })
            if (color.FieldCount(field) > 1 || color.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) return false;
        // Display P3 and CMYK are not silently reinterpreted as sRGB.
        if (color.HasField(12) && color.GetUnsigned(12) != 1) return false;
        // Keynote 15.4 writes this additional float for its opaque sRGB fills.
        // Native white and controlled blue/red exports qualify only the value 1;
        // its wider meaning is not defined by the independent Color schema.
        if (color.HasField(13) && (color.FieldCount(13) != 1
            || color.HasUnexpectedWireKind(13, IWorkWireKind.Fixed32)
            || color.GetFloat(13) != 1f || color.GetUnsigned(1) != 1
            || color.GetUnsigned(12) != 1 || color.HasField(11))) return false;
        return !color.HasField(1) || color.GetUnsigned(1) == (color.HasField(11) ? 3UL : 1UL);
    }

}
