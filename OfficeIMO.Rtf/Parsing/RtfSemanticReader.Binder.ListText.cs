using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private void ReadListText(RtfGroup group, CharacterState state, int depth) {
            CharacterState listTextState = state.Clone();
            listTextState.ListText = null;
            listTextState.PendingListTextAfterReset = null;
            RtfParagraph listText = ReadInlineParagraph(group, listTextState, depth);
            state.ListText = listText;
            state.PendingListTextAfterReset = listText;
        }
    }
}
