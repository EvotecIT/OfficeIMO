using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private void ReadNote(RtfGroup group, RtfNoteKind kind, CharacterState state, int depth) {
            RtfParagraph savedParagraph = _currentParagraph;
            RtfTable? savedTable = _currentTable;
            RtfTableRow? savedRow = _currentRow;
            RtfHeaderFooter? savedHeaderFooter = _currentHeaderFooter;
            RtfNote? savedNote = _currentNote;
            int savedCellIndex = _currentCellIndex;
            bool savedParagraphIsInTable = _currentParagraphIsInTable;

            var note = new RtfNote(kind);
            _currentNote = note;
            _currentParagraph = new RtfParagraph();
            _currentTable = null;
            _currentRow = null;
            _currentCellIndex = 0;
            _currentParagraphIsInTable = false;

            var childState = state.Clone();
            if (kind == RtfNoteKind.Annotation) {
                ReadAnnotationMetadata(group, note, childState.AnsiCodePage, childState.UnicodeSkipCount);
            }

            foreach (RtfNode node in group.Children) {
                switch (node) {
                    case RtfControlWord control when control.Name == group.Destination:
                        break;
                    case RtfControlWord control when kind == RtfNoteKind.Annotation && control.Name == "chatn":
                        break;
                    case RtfGroup childGroup when kind == RtfNoteKind.Annotation && IsAnnotationMetadataDestination(childGroup.Destination):
                        break;
                    case RtfGroup childGroup:
                        WalkGroup(childGroup, childState.Clone(), depth + 1, allowDestinationSkip: true);
                        break;
                    case RtfText text:
                        AppendText(ApplySkip(childState, RtfAnsiCodePage.DecodeText(childState.AnsiCodePage, text.Text)), childState);
                        break;
                    case RtfControlWord control:
                        ApplyControlWord(control, childState);
                        break;
                    case RtfControlSymbol symbol:
                        ApplyControlSymbol(symbol, childState);
                        break;
                }
            }

            FlushParagraphIfNeeded(force: false, childState);
            _document.AddParsedNote(note);

            _currentParagraph = savedParagraph;
            _currentTable = savedTable;
            _currentRow = savedRow;
            _currentHeaderFooter = savedHeaderFooter;
            _currentNote = savedNote;
            _currentCellIndex = savedCellIndex;
            _currentParagraphIsInTable = savedParagraphIsInTable;

            AttachNoteToCurrentRun(note);
        }

        private void AttachNoteToCurrentRun(RtfNote note) {
            IReadOnlyList<RtfRun> runs = _currentParagraph.Runs;
            if (runs.Count == 0) return;
            RtfRun run = runs[runs.Count - 1];
            if (run.Note == null) {
                run.Note = note;
            }
        }

        private static RtfNoteKind? TryGetNoteKind(string? destination) {
            switch (destination) {
                case "footnote":
                    return RtfNoteKind.Footnote;
                case "endnote":
                    return RtfNoteKind.Endnote;
                case "annotation":
                    return RtfNoteKind.Annotation;
                default:
                    return null;
            }
        }

        private static void ReadAnnotationMetadata(RtfGroup group, RtfNote note, int ansiCodePage, int unicodeSkipCount) {
            foreach (RtfGroup childGroup in group.Children.OfType<RtfGroup>()) {
                switch (childGroup.Destination) {
                    case "atnid":
                        note.Id = EmptyToNull(CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount).Trim());
                        break;
                    case "atnauthor":
                        note.Author = EmptyToNull(CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount).Trim());
                        break;
                    case "atntime":
                        note.Created = ReadInfoTimestamp(childGroup);
                        break;
                }
            }
        }

        private static bool IsAnnotationMetadataDestination(string? destination) =>
            destination == "atnid" || destination == "atnauthor" || destination == "atntime";
    }
}
