namespace OfficeIMO.IWork.Internal;

/// <summary>Reads effective table and cell-region stroke defaults through bounded native style inheritance.</summary>
internal static partial class IWorkTableBorderReader {
    private static readonly int[] DeprecatedFields = { 4, 5, 23, 24, 7, 8, 25, 26, 10, 11, 27, 28, 12, 29, 30, 31 };

    internal static IWorkTableBorderResolver Read(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        int rows, int columns, int headerRows, int headerColumns, int footerRows, bool ordinaryLayout) {
        var defaults = new IWorkTableBorderStyles(rows, columns, headerRows, headerColumns, footerRows);
        if (message.HasField(3)) {
            bool complete = true;
            var chain = ReadChain(source, model, message, 3, 6003, budget, references, ref complete);
            var strokes = new Dictionary<int, (IWorkArchiveRecord Owner, IWorkWireMessage Properties, int Field)>();
            var visibility = new Dictionary<int, (IWorkArchiveRecord Owner, IWorkWireMessage Properties)>();
            foreach (var item in chain.Reverse()) {
                source.CancellationToken.ThrowIfCancellationRequested();
                var properties = Properties(item.Record, item.Message, references, ref complete);
                if (properties == null) continue;
                for (int field = 46; field <= 61; field++) {
                    int selected = properties.HasField(field) ? field : DeprecatedFields[field - 46];
                    if (defaults.UsesStroke(field) && properties.HasField(selected)) strokes[field] = (item.Record, properties, selected);
                }
                foreach (int field in Enumerable.Range(46, 16).Where(defaults.UsesStroke)
                    .Select(IWorkTableBorderStyles.VisibilityField).Where(field => field > 0).Distinct())
                    if (properties.HasField(field)) visibility[field] = (item.Record, properties);
            }
            defaults.StructureIsComplete = complete;
            foreach (var pair in visibility) {
                var (owner, properties) = pair.Value;
                bool valid = IWorkStrokeReader.ValidScalar(properties, pair.Key, IWorkWireKind.Varint)
                    && properties.GetUnsigned(pair.Key) is 0 or 1;
                defaults.Visibility[pair.Key] = valid ? properties.GetUnsigned(pair.Key) == 1 : null;
                if (!valid) references.Declarations.Record(owner, "11/" + pair.Key, properties.FieldCount(pair.Key), IWorkSourceDeclarationIssueKind.InvalidValue);
            }
            foreach (var pair in strokes) {
                int flag = IWorkTableBorderStyles.VisibilityField(pair.Key);
                // Hidden inherited strokes do not affect the rendered edge.
                if (defaults.Visibility.TryGetValue(flag, out bool? visible) && visible == false) continue;
                var (owner, properties, field) = pair.Value;
                defaults.Strokes[pair.Key] = IWorkStrokeReader.Read(owner, properties, field, "11/" + field, references, field < 46);
            }
        }
        for (int field = 18; field <= 21; field++) {
            bool applicable = field switch {
                18 => rows > headerRows + footerRows && columns > headerColumns,
                19 => headerRows > 0,
                20 => rows > headerRows + footerRows && headerColumns > 0,
                _ => footerRows > 0
            };
            if (!applicable || !message.HasField(field)) continue;
            bool complete = true;
            var role = new IWorkCellBorderStyle();
            foreach (var item in ReadChain(source, model, message, field, 6004, budget, references, ref complete).Reverse()) {
                source.CancellationToken.ThrowIfCancellationRequested();
                var properties = Properties(item.Record, item.Message, references, ref complete);
                if (properties != null) role.Read(item.Record, properties, references);
            }
            if (!complete) role.Invalidate();
            defaults.CellRoles[field] = role;
        }
        if (defaults.HasDeclarations && (headerRows > 1 || headerColumns > 1 || footerRows > 1)
            || !defaults.RoleEdgesAreUnambiguous()) {
            defaults.StructureIsComplete = false;
            references.Declarations.Record(model, "$table-border-regions", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        }
        var overrides = ReadOverrides(source, model, message, budget, references, rows, columns);
        if (!ordinaryLayout) { defaults.StructureIsComplete = false; overrides.RejectLocations(); }
        return new IWorkTableBorderResolver(defaults, overrides);
    }

    private static IReadOnlyList<(IWorkArchiveRecord Record, IWorkWireMessage Message)> ReadChain(
        IWorkSourceDocument source, IWorkArchiveRecord owner, IWorkWireMessage message, int field, uint type,
        IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references, ref bool complete) {
        var style = references.ReadOne(owner, message, field, allowedType: candidate => candidate == type);
        complete = message.FieldCount(field) == 1 && !message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
            && style?.MessageType == type;
        if (!complete) {
            references.Declarations.Record(owner, field.ToString(System.Globalization.CultureInfo.InvariantCulture), message.FieldCount(field), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            return Array.Empty<(IWorkArchiveRecord, IWorkWireMessage)>();
        }
        return IWorkStyleReader.ReadChain(source.Index, style!.Identifier, budget.MaximumTextStyleInheritanceDepth,
            candidate => candidate == type, tolerateStyleDepth: true, references, ref complete);
    }

    private static IWorkWireMessage? Properties(IWorkArchiveRecord owner, IWorkWireMessage message,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        var properties = IWorkObjectIndex.TryGetMessage(message, 11, out bool malformed);
        if (malformed || message.FieldCount(11) > 1 || message.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
            || message.HasField(11) && properties == null) {
            references.Declarations.Record(owner, "11", message.FieldCount(11)); complete = false;
        }
        return properties;
    }
}
