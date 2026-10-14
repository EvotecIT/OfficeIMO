namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkPagesArchiveBuilder {
    private void AddPageTextStorage(ulong storageId, IWorkTextRunData? run, bool hasPageText) {
        _cancellationToken.ThrowIfCancellationRequested();
        ulong paragraphId = HeaderParagraphId, characterId = HeaderCharacterId;
        if (run != null) {
            paragraphId = Allocate(); characterId = Allocate();
            Add(paragraphId, 2022, P().Message(1, Style(paragraphId, "OfficeIMO Page Text"))
                .Message(11, Character(run)).Message(12, ParagraphProperties(IWorkPagesParagraphLayout.Default)).UInt(10, 10));
            Add(characterId, 2021, P().Message(1, Style(characterId, "OfficeIMO Page Text"))
                .Message(11, Character(run)).UInt(10, 8));
        }
        // Native headers/footers have no body paragraph terminator. Empty positions
        // remain present, but have no text to become extra destination paragraphs.
        // Retain the established empty graph exactly when page text is disabled.
        string text = run?.Text ?? (hasPageText ? string.Empty : "\n");
        Add(storageId, 2001, P().UInt(1, 1).Reference(2, StylesheetId).String(3, text).Bool(10, true)
            .Message(5, Attribute(paragraphId)).Message(7, Attribute(ListId)).Message(8, Attribute(characterId))
            .Message(12, Attribute(ColumnId)).Message(28, P().Message(1, P().UInt(1, 0)))
            .Message(6, EmptyParagraphData()).Message(14, EmptyParagraphData()).Message(24, EmptyParagraphData()));
    }
}
