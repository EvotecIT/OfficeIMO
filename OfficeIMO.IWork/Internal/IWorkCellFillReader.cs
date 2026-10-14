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
                && color is { Alpha: byte.MaxValue }
                && IWorkColorReader.IsSupportedColor(IWorkObjectIndex.TryGetMessage(declaration, 1)!);
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

}
