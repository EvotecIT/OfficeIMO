namespace OfficeIMO.IWork.Internal;

/// <summary>Decodes the supported cell FillArchive subset for selected and role styles.</summary>
internal static class IWorkCellFillReader {
    internal static void Read(IWorkArchiveRecord owner, IWorkWireMessage properties,
        IWorkSourceReferenceIssueCollector references, ref IWorkCellFill? fill, ref bool complete) {
        if (!properties.HasField(1)) return;
        IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, 1, out bool malformed);
        bool valid = !malformed && properties.FieldCount(1) == 1
            && !properties.HasUnexpectedWireKind(1, IWorkWireKind.Bytes) && declaration != null;
        IWorkColor? color = null;
        if (valid && declaration!.TotalFieldCount > 0) {
            valid = declaration.TotalFieldCount == 1 && declaration.FieldCount(1) == 1
                && IWorkColorReader.TryRead(declaration, 1, out color, ref valid)
                && color is { Alpha: byte.MaxValue } && IsSupportedColor(declaration);
        }
        if (!valid) {
            references.Declarations.Record(owner, "11/1", properties.FieldCount(1),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        // An empty FillArchive clears the parent rather than inheriting its color.
        fill = new IWorkCellFill(color);
    }

    private static bool IsSupportedColor(IWorkWireMessage fill) {
        IWorkWireMessage color = IWorkObjectIndex.TryGetMessage(fill, 1)!;
        int[] fields = { 1, 3, 4, 5, 6, 11, 12 };
        if (color.TotalFieldCount != fields.Sum(color.FieldCount)) return false;
        foreach (int field in new[] { 1, 12 })
            if (color.FieldCount(field) > 1 || color.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) return false;
        // Display P3 and CMYK are not silently reinterpreted as sRGB.
        if (color.HasField(12) && color.GetUnsigned(12) != 1) return false;
        return !color.HasField(1) || color.GetUnsigned(1) == (color.HasField(11) ? 3UL : 1UL);
    }

}
