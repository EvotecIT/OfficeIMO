namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableBorderReader {
    private static IWorkTableBorderOverrides ReadOverrides(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        int rows, int columns) {
        var result = new IWorkTableBorderOverrides();
        if (!message.HasField(49)) return result;
        IWorkArchiveRecord? sidecar = references.ReadOne(model, message, 49, allowedType: type => type == 6305);
        if (sidecar?.MessageType != 6305 || message.FieldCount(49) != 1
            || message.HasUnexpectedWireKind(49, IWorkWireKind.Bytes)) { Reject(model, message, "49", 49); return result; }
        try {
            var payload = source.Index.Message(sidecar);
            bool valid = payload.TotalFieldCount == Enumerable.Range(1, 7).Sum(payload.FieldCount)
                && Enumerable.Range(1, 3).All(field => IWorkStrokeReader.ValidScalar(payload, field, IWorkWireKind.Varint))
                && payload.GetUnsigned(1).GetValueOrDefault() <= uint.MaxValue
                && payload.GetUnsigned(2) == (ulong)columns && payload.GetUnsigned(3) == (ulong)rows;
            if (!valid) { Reject(sidecar, payload, "$", null); return result; }
            uint maxOrder = (uint)payload.GetUnsigned(1).GetValueOrDefault();
            budget.AddTableBorderEntries(Enumerable.Range(4, 4).Sum(payload.FieldCount));
            var seen = new HashSet<(int Field, ulong Identifier)>();
            for (int field = 4; field <= 7; field++) {
                source.CancellationToken.ThrowIfCancellationRequested();
                var layers = references.ReadAll(sidecar, payload, field, out int unresolved, allowedType: type => type == 6306);
                if (unresolved > 0 || layers.Count != payload.FieldCount(field)) { Reject(sidecar, payload, field.ToString(), field); continue; }
                foreach (var layer in layers) {
                    source.CancellationToken.ThrowIfCancellationRequested();
                    if (layer.MessageType != 6306 || !seen.Add((field, layer.Identifier))) { Reject(sidecar, payload, field.ToString(), field); continue; }
                    ReadLayer(layer, field, maxOrder);
                }
            }
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            result.RejectLocations(); references.Declarations.Record(sidecar, "$", null);
        }
        result.Seal(source.CancellationToken);
        if (!result.FullyReconstructed) references.Declarations.Record(sidecar, "$", null, IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        return result;

        void ReadLayer(IWorkArchiveRecord layer, int field, uint maxOrder) {
            var payload = source.Index.Message(layer);
            bool horizontal = field >= 6;
            int lineCount = horizontal ? rows : columns, positionCount = horizontal ? columns : rows;
            ulong? index = payload.GetUnsigned(1);
            bool valid = payload.TotalFieldCount == payload.FieldCount(1) + payload.FieldCount(2)
                && IWorkStrokeReader.ValidScalar(payload, 1, IWorkWireKind.Varint) && index.HasValue && index < (ulong)lineCount
                && !payload.HasUnexpectedWireKind(2, IWorkWireKind.Bytes);
            budget.AddTableBorderEntries(payload.FieldCount(2));
            var runs = IWorkObjectIndex.TryGetMessages(payload, 2, out bool malformed);
            if (!valid || malformed || runs.Count != payload.FieldCount(2)) { Reject(layer, payload, "$", null); return; }
            int line = checked((int)index!.Value + (field is 5 or 7 ? 1 : 0));
            int position = 0;
            foreach (var run in runs) {
                source.CancellationToken.ThrowIfCancellationRequested(); position++;
                ulong? origin = run.GetUnsigned(1), length = run.GetUnsigned(2);
                ulong order = run.GetUnsigned(4).GetValueOrDefault();
                valid = run.TotalFieldCount == Enumerable.Range(1, 4).Sum(run.FieldCount)
                    && new[] { 1, 2, 4 }.All(scalar => IWorkStrokeReader.ValidScalar(run, scalar, IWorkWireKind.Varint))
                    && origin.HasValue && origin < (ulong)positionCount && length > 0
                    && length <= (ulong)positionCount - origin && order <= maxOrder;
                string path = "2[" + position + "]";
                if (!valid) { Reject(layer, run, path, null); continue; }
                var stroke = IWorkStrokeReader.Read(layer, run, 3, path + "/3", references);
                result.Add(horizontal, line, (int)origin!.Value, (int)length!.Value, (uint)order, stroke);
            }
        }

        void Reject(IWorkArchiveRecord owner, IWorkWireMessage payload, string path, int? field) {
            result.RejectLocations(); references.Declarations.Record(owner, path,
                field.HasValue ? payload.FieldCount(field.Value) : null, IWorkSourceDeclarationIssueKind.InvalidValue);
        }
    }
}
