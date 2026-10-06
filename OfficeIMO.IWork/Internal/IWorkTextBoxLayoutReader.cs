namespace OfficeIMO.IWork.Internal;

/// <summary>Reads selected horizontal, single-column text-frame styles.</summary>
internal static class IWorkTextBoxLayoutReader {
    internal static IWorkTextBoxLayout? Read(IWorkObjectIndex index, IWorkArchiveRecord drawable,
        IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references, out bool complete) {
        IWorkWireMessage? shape = IWorkDrawingReader.ShapeMessage(index, drawable, out complete);
        if (shape == null) return null;
        IWorkWireMessage archive = index.Message(drawable);
        IWorkWireMessage? info = drawable.MessageType == 7
            ? IWorkObjectIndex.TryGetMessage(archive, 1) : archive;
        if (info?.HasField(3) == true) {
            Reject(drawable, drawable.MessageType == 7 ? "1/3" : "3", info.FieldCount(3), references, ref complete);
        }
        // The existing direct-storage path also accepts a legacy nested storage
        // envelope. It is a shape-style path only with the native drawable super.
        if (!shape.HasField(1)) return null;
        string shapePath = drawable.MessageType == 7 ? "1/1" : "1";
        IWorkArchiveRecord? style = references.ReadOne(drawable, shape, 2, shapePath + "/2", type => type == 2025);
        if (!shape.HasField(2)) return null;
        if (style?.MessageType != 2025) { complete = false; return null; }
        IReadOnlyList<(IWorkArchiveRecord Record, IWorkWireMessage Message)> chain = IWorkStyleReader.ReadChain(
            index, style.Identifier, budget.MaximumTextStyleInheritanceDepth, type => type == 2025,
            false, references, ref complete, superArchiveDepth: 2);
        bool? shrink = null;
        bool singleColumn = false;
        IWorkTextVerticalAlignment? alignment = null;
        double? left = null, top = null, right = null, bottom = null;
        for (int level = chain.Count - 1; level >= 0; level--) {
            var item = chain[level];
            IWorkWireMessage? properties = Child(item.Message, 11, item.Record, "11", references, ref complete);
            if (properties == null) continue;
            bool? fit = Boolean(properties, 1, item.Record, references, ref complete);
            if (fit.HasValue) shrink = fit;
            if (properties.HasField(2)) {
                ulong? value = Unsigned(properties, 2, item.Record, "11/2", references, ref complete);
                if (value <= 2) alignment = (IWorkTextVerticalAlignment)value!.Value;
                else { Reject(item.Record, "11/2", properties.FieldCount(2), references, ref complete); }
            }
            // Vertical text and multiple columns need separate destination qualification.
            foreach (int field in new[] { 8, 11 }) {
                if (Boolean(properties, field, item.Record, references, ref complete) == true)
                    Reject(item.Record, "11/" + field, properties.FieldCount(field), references, ref complete);
            }
            if (Boolean(properties, 3, item.Record, references, ref complete) == true)
                Reject(item.Record, "11/3", properties.FieldCount(3), references, ref complete);
            IWorkWireMessage? columns = Child(properties, 4, item.Record, "11/4", references, ref complete);
            if (columns != null) {
                IWorkWireMessage? equal = Child(columns, 1, item.Record, "11/4/1", references, ref complete);
                if (columns.HasField(2) || equal == null
                    || Unsigned(equal, 1, item.Record, "11/4/1/1", references, ref complete) != 1)
                    Reject(item.Record, "11/4", properties.FieldCount(4), references, ref complete);
                else singleColumn = true;
            }
            if (Boolean(properties, 5, item.Record, references, ref complete) == true)
                Reject(item.Record, "11/5", properties.FieldCount(5), references, ref complete);
            IWorkWireMessage? padding = Child(properties, 6, item.Record, "11/6", references, ref complete);
            if (padding != null) {
                if (padding.TotalFieldCount != Enumerable.Range(1, 4).Sum(padding.FieldCount))
                    Reject(item.Record, "11/6", properties.FieldCount(6), references, ref complete);
                // A present PaddingArchive replaces the whole inherited property. Apple omits zero-valued sides.
                left = Float(padding, 1, item.Record, references, ref complete);
                top = Float(padding, 2, item.Record, references, ref complete);
                right = Float(padding, 3, item.Record, references, ref complete);
                bottom = Float(padding, 4, item.Record, references, ref complete);
            }
            if (properties.HasField(7)
                && Unsigned(properties, 7, item.Record, "11/7", references, ref complete) != 0)
                Reject(item.Record, "11/7", properties.FieldCount(7), references, ref complete);
            if (Boolean(properties, 9, item.Record, references, ref complete) == true
                && properties.HasField(10))
                Reject(item.Record, "11/10", properties.FieldCount(10), references, ref complete);
            if (properties.HasField(10)) {
                IWorkArchiveRecord? paragraph = references.ReadOne(item.Record, properties, 10, "11/10", type => type == 2022);
                if (paragraph?.MessageType != 2022) complete = false;
            }
        }
        return complete ? new IWorkTextBoxLayout(shrink, alignment, left, top, right, bottom, singleColumn) : null;
    }

    // Native fixed-frame probes include wrapped paragraphs with keep-lines on/off.
    // Connected flows, unknown frame layouts, notes and cells retain the pagination gate.
    internal static bool HasUnsupportedPagination(IWorkParagraphStyle style, IWorkTextBox? frame = null) =>
        style.PageBreakBefore == true || style.KeepWithNext == true
        || style.KeepLinesTogether == true && !(frame?.Layout?.SupportsFixedFramePagination == true
            && IWorkDrawingReader.HasPositiveSize(frame.Geometry));

    private static IWorkWireMessage? Child(IWorkWireMessage owner, int field, IWorkArchiveRecord record,
        string path, IWorkSourceReferenceIssueCollector references, ref bool complete) {
        IWorkWireMessage? result = IWorkObjectIndex.TryGetMessage(owner, field, out bool malformed);
        if (malformed || owner.HasField(field) && result == null) {
            references.Declarations.Record(record, path, owner.FieldCount(field)); complete = false;
        }
        return result;
    }

    private static ulong? Unsigned(IWorkWireMessage owner, int field, IWorkArchiveRecord record,
        string path, IWorkSourceReferenceIssueCollector references, ref bool complete) {
        ulong? value = owner.GetUnsigned(field);
        if (owner.FieldCount(field) != 1 || owner.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) {
            references.Declarations.Record(record, path, owner.FieldCount(field)); complete = false; return null;
        }
        return value;
    }

    private static bool? Boolean(IWorkWireMessage owner, int field, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        if (!owner.HasField(field)) return null;
        ulong? value = Unsigned(owner, field, record, "11/" + field, references, ref complete);
        if (value <= 1) return value == 1;
        references.Declarations.Record(record, "11/" + field, owner.FieldCount(field)); complete = false; return null;
    }

    private static double? Float(IWorkWireMessage owner, int field, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        if (!owner.HasField(field)) return 0;
        float? value = owner.GetFloat(field);
        if (owner.FieldCount(field) != 1 || owner.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
            || !value.HasValue || float.IsNaN(value.Value) || float.IsInfinity(value.Value) || value < 0) {
            references.Declarations.Record(record, "11/6/" + field, owner.FieldCount(field)); complete = false; return null;
        }
        return value;
    }

    private static void Reject(IWorkArchiveRecord record, string path, int count,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        references.Declarations.Record(record, path, count, IWorkSourceDeclarationIssueKind.UnsupportedField);
        complete = false;
    }
}
