namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    private static string RenderColonTract(IWorkWireMessage node, int row, int column, int maximumCharacters,
        IReadOnlyDictionary<Guid, IWorkFormulaTableBinding>? tableQualifiers, IWorkFormulaTableBinding? owningTable,
        ref bool complete, ref bool boundedBodyRanges, ref bool needsLocalBodyBinding) {
        IWorkWireMessage? tract = IWorkObjectIndex.TryGetMessage(node, 40, out bool malformedTract);
        if (malformedTract || tract == null) { complete = false; return "#REF!"; }
        int? firstColumn, lastColumn, firstRow, lastRow;
        bool firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute;
        if (node.HasField(33)) {
            IWorkWireMessage? sticky = IWorkObjectIndex.TryGetMessage(node, 33, out bool malformedSticky);
            if (malformedSticky || sticky == null || sticky.TotalFieldCount != Enumerable.Range(1, 4).Sum(sticky.FieldCount)
                || !TrySticky(sticky, 1, out firstRowAbsolute) || !TrySticky(sticky, 2, out firstColumnAbsolute)
                || !TrySticky(sticky, 3, out lastRowAbsolute) || !TrySticky(sticky, 4, out lastColumnAbsolute)
                || !TryEndpointRange(tract, 1, 3, column, firstColumnAbsolute, lastColumnAbsolute,
                    32767, out firstColumn, out lastColumn)
                || !TryEndpointRange(tract, 2, 4, row, firstRowAbsolute, lastRowAbsolute,
                    int.MaxValue, out firstRow, out lastRow)) { complete = false; return "#REF!"; }
        } else {
            if (!TryRange(tract, 3, 1, column, out int legacyFirstColumn, out int legacyLastColumn, out firstColumnAbsolute)
                || !TryRange(tract, 4, 2, row, out int legacyFirstRow, out int legacyLastRow, out firstRowAbsolute)) {
                complete = false; return "#REF!";
            }
            firstColumn = legacyFirstColumn; lastColumn = legacyLastColumn;
            firstRow = legacyFirstRow; lastRow = legacyLastRow;
            lastColumnAbsolute = firstColumnAbsolute; lastRowAbsolute = firstRowAbsolute;
        }
        if (!firstColumn.HasValue || !firstRow.HasValue) needsLocalBodyBinding = true;
        return BindTableReference(node, firstColumn, lastColumn, firstRow, lastRow,
            firstColumnAbsolute, lastColumnAbsolute, firstRowAbsolute, lastRowAbsolute, maximumCharacters,
            tableQualifiers, owningTable, ref complete, ref boundedBodyRanges);
    }

    private static bool TrySticky(IWorkWireMessage sticky, int field, out bool absolute) {
        absolute = sticky.GetUnsigned(field) == 1;
        return sticky.FieldCount(field) <= 1 && !sticky.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
            && sticky.GetUnsigned(field).GetValueOrDefault() <= 1;
    }

    private static bool TryEndpointRange(IWorkWireMessage tract, int relativeField, int absoluteField, int origin,
        bool firstAbsolute, bool lastAbsolute, int wholeAxisSentinel, out int? first, out int? last) {
        first = last = null;
        if (!TryCoordinateRange(tract, relativeField, relative: true, out long? relativeBegin, out long? relativeEnd)
            || !TryCoordinateRange(tract, absoluteField, relative: false, out long? absoluteBegin, out long? absoluteEnd)) return false;
        // An absent relative axis with the exact native absolute sentinel spans the table body.
        if (!relativeBegin.HasValue && absoluteBegin == wholeAxisSentinel && absoluteEnd == wholeAxisSentinel) return true;
        long? begin = firstAbsolute ? absoluteBegin : relativeBegin + origin;
        long? end = lastAbsolute ? absoluteEnd : relativeEnd + origin;
        if (!begin.HasValue || !end.HasValue || begin < 0 || end < begin || end > int.MaxValue) return false;
        first = (int)begin.Value; last = (int)end.Value;
        return true;
    }

    private static bool TryCoordinateRange(IWorkWireMessage tract, int field, bool relative,
        out long? begin, out long? end) {
        begin = end = null;
        if (!tract.HasField(field)) return true;
        IWorkWireMessage? range = IWorkObjectIndex.TryGetMessage(tract, field, out bool malformed);
        if (malformed || range == null || range.TotalFieldCount != range.FieldCount(1) + range.FieldCount(2) || range.FieldCount(1) != 1
            || range.FieldCount(2) > 1 || range.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || range.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            || range.GetUnsigned(1) is not ulong rawBegin) return false;
        ulong rawEnd = range.GetUnsigned(2) ?? rawBegin;
        if (relative) {
            if (!TrySignedCoordinate(rawBegin, out int signedBegin) || !TrySignedCoordinate(rawEnd, out int signedEnd)) return false;
            begin = signedBegin; end = signedEnd;
        } else {
            if (rawBegin > int.MaxValue || rawEnd > int.MaxValue) return false;
            begin = (long)rawBegin; end = (long)rawEnd;
        }
        return true;
    }

    // Protobuf int32 negative values may use the sign-extended ten-byte varint;
    // the existing compact uint32 representation remains accepted as well.
    private static bool TrySignedCoordinate(ulong raw, out int value) {
        value = unchecked((int)(uint)raw);
        return raw <= uint.MaxValue || raw >= unchecked((ulong)(long)int.MinValue);
    }

    private static bool TryRange(IWorkWireMessage tract, int absoluteField, int relativeField, int origin,
        out int first, out int last, out bool absolute) {
        if (TryAbsoluteRange(tract, absoluteField, out first, out last)) {
            absolute = true;
            return true;
        }
        absolute = false;
        IReadOnlyList<IWorkWireMessage> ranges = IWorkObjectIndex.TryGetMessages(tract, relativeField, out bool malformed);
        if (malformed || ranges.Count != 1
            || ranges[0].FieldCount(1) != 1 || ranges[0].FieldCount(2) > 1
            || ranges[0].HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || ranges[0].HasUnexpectedWireKind(2, IWorkWireKind.Varint)) {
            first = last = 0;
            return false;
        }
        ulong rawBegin = ranges[0].GetUnsigned(1) ?? 0;
        ulong rawEnd = ranges[0].GetUnsigned(2) ?? rawBegin;
        if (!TrySignedCoordinate(rawBegin, out int begin) || !TrySignedCoordinate(rawEnd, out int end)) {
            first = last = 0;
            return false;
        }
        long resolvedFirst = (long)origin + begin;
        long resolvedLast = (long)origin + end;
        if (resolvedFirst < 0 || resolvedFirst > int.MaxValue
            || resolvedLast < resolvedFirst || resolvedLast > int.MaxValue) {
            first = last = 0;
            return false;
        }
        first = (int)resolvedFirst;
        last = (int)resolvedLast;
        return first >= 0 && last >= first;
    }

}
