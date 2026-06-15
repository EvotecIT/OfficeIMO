using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private void ReadBookmarkMarker(RtfGroup group, RtfBookmarkMarkerKind kind, CharacterState state) {
            string name = CollectPlainText(group, state.AnsiCodePage, state.UnicodeSkipCount).Trim();
            if (name.Length == 0) {
                return;
            }

            _currentParagraph.AddBookmarkMarker(new RtfBookmarkMarker(kind, name));
        }

        private static RtfBookmarkMarkerKind? TryGetBookmarkMarkerKind(string? destination) {
            switch (destination) {
                case "bkmkstart":
                    return RtfBookmarkMarkerKind.Start;
                case "bkmkend":
                    return RtfBookmarkMarkerKind.End;
                default:
                    return null;
            }
        }
    }
}
