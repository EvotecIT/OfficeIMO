using System.Globalization;

namespace OfficeIMO.IWork.Internal;

/// <summary>Observes only reference fields selected by a projection; it does not walk inactive records.</summary>
internal sealed class IWorkSourceReferenceIssueCollector(IWorkSourceDocument source) {
    private readonly List<IWorkSourceReferenceIssue> _issues = new();
    private readonly HashSet<(ulong Owner, string Path)> _recordedFields = new();
    private int _inspectedReferenceCount;

    internal IReadOnlyList<IWorkSourceReferenceIssue> Issues => _issues;
    internal IWorkSourceDeclarationIssueCollector Declarations { get; } = new(source);

    internal IWorkArchiveRecord? ReadOne(IWorkArchiveRecord owner, IWorkWireMessage message,
        int field, string? path = null, Func<uint, bool>? allowedType = null) {
        IWorkArchiveRecord? result = source.Index.Dereference(message, field);
        if (result != null && allowedType != null && !allowedType(result.MessageType)) {
            Record(owner, message, field, path, rejectedSet: false, allowedType);
            return null;
        }
        if (result == null && message.HasField(field))
            Record(owner, message, field, path, rejectedSet: message.FieldCount(field) > 1);
        return result;
    }

    internal IReadOnlyList<IWorkArchiveRecord> ReadAll(IWorkArchiveRecord owner, IWorkWireMessage message,
        int field, out int unresolved, string? path = null) {
        IReadOnlyList<IWorkArchiveRecord> result = source.Index.DereferenceAll(message, field, out unresolved,
            out bool rejectedSet);
        if (unresolved > 0) {
            // A malformed repeated field is rejected as a whole by the existing reader. Retain
            // evidence for readable siblings too, without calling their existing targets missing.
            Record(owner, message, field, path, rejectedSet);
        }
        return result;
    }

    private void Record(IWorkArchiveRecord owner, IWorkWireMessage message, int field,
        string? path, bool rejectedSet, Func<uint, bool>? allowedType = null) {
        string fieldPath = path ?? field.ToString(CultureInfo.InvariantCulture);
        var fieldKey = (owner.Identifier, fieldPath);
        // Shared text/template archives may be selected more than once. Their physical
        // reference occurrences are source evidence, independent of the number of uses.
        if (_recordedFields.Contains(fieldKey)) return;
        int declared = message.FieldCount(field);
        // Charge the entire selected field before parsing/materializing evidence. This conservative
        // bound also limits the inspection work when only some occurrences actually fail.
        if (declared > source.Options.MaximumSourceReferenceIssues - _inspectedReferenceCount)
            throw new InvalidDataException($"iWork source reference issues exceed the configured limit of {source.Options.MaximumSourceReferenceIssues}.");
        _inspectedReferenceCount += declared;
        _recordedFields.Add(fieldKey);
        var identity = new IWorkObjectIdentity(owner);
        int position = 0;
        foreach (IWorkWireValue value in message.EnumerateValues(field)) {
            source.CancellationToken.ThrowIfCancellationRequested();
            position++;
            ulong? identifier = null;
            bool malformed = value.Kind != IWorkWireKind.Bytes || value.Bytes == null;
            if (!malformed) {
                try {
                    IWorkWireMessage reference = message.ParseNestedMessage(value.Bytes!);
                    malformed = reference.FieldCount(1) != 1
                        || reference.HasUnexpectedWireKind(1, IWorkWireKind.Varint);
                    if (!malformed) identifier = reference.GetUnsigned(1);
                } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
                    malformed = true;
                }
            }
            IWorkArchiveRecord? target = identifier is { } id ? source.Index.Find(id) : null;
            IWorkSourceReferenceIssueKind? kind = malformed ? IWorkSourceReferenceIssueKind.MalformedReference
                : rejectedSet ? IWorkSourceReferenceIssueKind.RejectedReferenceSet
                : identifier.HasValue && target == null ? IWorkSourceReferenceIssueKind.MissingTarget
                : target != null && allowedType != null && !allowedType(target.MessageType)
                    ? IWorkSourceReferenceIssueKind.UnexpectedTargetType
                : null;
            if (kind is { } issueKind)
                _issues.Add(new IWorkSourceReferenceIssue(identity, fieldPath, position, identifier, issueKind));
        }
    }
}
