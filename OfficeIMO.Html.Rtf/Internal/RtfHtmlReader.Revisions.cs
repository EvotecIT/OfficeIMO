using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyParagraphRevisionAttributes(IElement token) {
            int? revisionSaveId = ReadIntegerAttribute(token, "data-officeimo-rtf-pararsid");
            if (revisionSaveId.HasValue) {
                EnsureParagraph().RevisionSaveId = revisionSaveId.Value;
            }
        }

        private void PushRevisionScope(IElement token) {
            RtfRevisionKind? kind = ReadRevisionKind(token);
            bool hasMetadata = kind.HasValue ||
                               HasAttribute(token, "data-officeimo-rtf-revision-author") ||
                               HasAttribute(token, "data-officeimo-rtf-revision-author-index") ||
                               HasAttribute(token, "data-officeimo-rtf-revision-timestamp") ||
                               HasAttribute(token, "data-officeimo-rtf-charrsid") ||
                               HasAttribute(token, "data-officeimo-rtf-insrsid") ||
                               HasAttribute(token, "data-officeimo-rtf-delrsid");
            if (!hasMetadata) {
                return;
            }

            _revisions.Push(new RtfRevisionScope(
                token.LocalName,
                kind ?? RtfRevisionKind.None,
                ReadRevisionAuthorIndex(token),
                ReadIntegerAttribute(token, "data-officeimo-rtf-revision-timestamp"),
                ReadIntegerAttribute(token, "data-officeimo-rtf-charrsid"),
                ReadIntegerAttribute(token, "data-officeimo-rtf-insrsid"),
                ReadIntegerAttribute(token, "data-officeimo-rtf-delrsid")));
        }

        private void PopRevisionScope(string name) {
            if (_revisions.Count == 0) {
                return;
            }

            var deferred = new List<RtfRevisionScope>();
            while (_revisions.Count > 0) {
                RtfRevisionScope scope = _revisions.Pop();
                if (string.Equals(scope.Name, name, StringComparison.OrdinalIgnoreCase)) {
                    break;
                }

                deferred.Add(scope);
            }

            for (int index = deferred.Count - 1; index >= 0; index--) {
                _revisions.Push(deferred[index]);
            }
        }

        private void ApplyRevision(RtfRun run) {
            foreach (RtfRevisionScope scope in _revisions) {
                run.RevisionKind = scope.Kind;
                run.RevisionAuthorIndex = scope.AuthorIndex;
                run.RevisionTimestampValue = scope.TimestampValue;
                run.CharacterRevisionSaveId = scope.CharacterRevisionSaveId;
                run.InsertionRevisionSaveId = scope.InsertionRevisionSaveId;
                run.DeletionRevisionSaveId = scope.DeletionRevisionSaveId;
                return;
            }
        }

        private RtfRevisionKind? ReadRevisionKind(IElement token) {
            string? value = GetAttribute(token, "data-officeimo-rtf-revision");
            if (!string.IsNullOrWhiteSpace(value)) {
                switch (value!.Trim().ToLowerInvariant()) {
                    case "inserted":
                    case "insert":
                    case "ins":
                        return RtfRevisionKind.Inserted;
                    case "deleted":
                    case "delete":
                    case "del":
                        return RtfRevisionKind.Deleted;
                    case "none":
                        return RtfRevisionKind.None;
                }
            }

            if (string.Equals(token.LocalName, "ins", StringComparison.OrdinalIgnoreCase)) {
                return RtfRevisionKind.Inserted;
            }

            return string.Equals(token.LocalName, "del", StringComparison.OrdinalIgnoreCase)
                ? RtfRevisionKind.Deleted
                : null;
        }

        private int? ReadRevisionAuthorIndex(IElement token) {
            int? explicitIndex = ReadIntegerAttribute(token, "data-officeimo-rtf-revision-author-index");
            string? author = GetAttribute(token, "data-officeimo-rtf-revision-author");
            if (string.IsNullOrWhiteSpace(author)) {
                return explicitIndex;
            }

            int existingIndex = FindRevisionAuthor(author!);
            if (existingIndex < 0) {
                existingIndex = _document.AddRevisionAuthor(author!);
            }

            return explicitIndex ?? existingIndex;
        }

        private int FindRevisionAuthor(string author) {
            for (int index = 0; index < _document.RevisionAuthors.Count; index++) {
                if (string.Equals(_document.RevisionAuthors[index].Name, author, StringComparison.Ordinal)) {
                    return index;
                }
            }

            return -1;
        }

        private static int? ReadIntegerAttribute(IElement token, string name) {
            string? value = GetAttribute(token, name);
            if (string.IsNullOrWhiteSpace(value) ||
                !int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ||
                parsed < 0) {
                return null;
            }

            return parsed;
        }

        private static bool HasAttribute(IElement token, string name) {
            return token.HasAttribute(name);
        }
    }

    private sealed class RtfRevisionScope {
        internal RtfRevisionScope(
            string name,
            RtfRevisionKind kind,
            int? authorIndex,
            int? timestampValue,
            int? characterRevisionSaveId,
            int? insertionRevisionSaveId,
            int? deletionRevisionSaveId) {
            Name = name;
            Kind = kind;
            AuthorIndex = authorIndex;
            TimestampValue = timestampValue;
            CharacterRevisionSaveId = characterRevisionSaveId;
            InsertionRevisionSaveId = insertionRevisionSaveId;
            DeletionRevisionSaveId = deletionRevisionSaveId;
        }

        internal string Name { get; }

        internal RtfRevisionKind Kind { get; }

        internal int? AuthorIndex { get; }

        internal int? TimestampValue { get; }

        internal int? CharacterRevisionSaveId { get; }

        internal int? InsertionRevisionSaveId { get; }

        internal int? DeletionRevisionSaveId { get; }
    }
}
