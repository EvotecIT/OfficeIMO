namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private static int ClampTextRunBoundary(uint boundary, int textLength) {
        if (boundary > int.MaxValue) {
            throw new OneNoteFormatException(
                "ONENOTE_TEXT_RUN_BOUNDARY",
                "A rich-text run boundary exceeds the supported text length.");
        }
        return Math.Min(textLength, (int)boundary);
    }
}
