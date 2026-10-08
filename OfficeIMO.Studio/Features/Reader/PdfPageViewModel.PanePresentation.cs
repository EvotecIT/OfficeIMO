namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageViewModel {
    /// <summary>Mirrors canonical edit/search/field state without sharing viewport attachment, images or zoom.</summary>
    internal void MirrorInteractionState(PdfPageViewModel source, bool active) {
        EditorTool = source.EditorTool;
        SelectionMode = source.SelectionMode;
        IsNightMode = source.IsNightMode;
        SelectedObject = source.SelectedObject;
        InlineTextDraft = active ? source.InlineTextDraft : null;
        SelectedAnnotations = source.SelectedAnnotations;
        CommentAnchorObjectNumber = source.CommentAnchorObjectNumber;
        PendingRedactionAreas = source.PendingRedactionAreas;
        SearchHighlights = source.SearchHighlights;
        ActiveSearchHighlight = source.ActiveSearchHighlight;
        ActiveSearchHighlights = source.ActiveSearchHighlights;
        FormAnchorFieldName = source.FormAnchorFieldName;
        if (!active) FocusInlineFormEditorRequested = false;
        ShowInlineFormFields(source._inlineFormFields, source.InlineFormField, active && source.FocusInlineFormEditorRequested);
        // Changing the field clears the local widget anchor, so mirror its identity afterwards.
        FormAnchorObjectNumber = source.FormAnchorObjectNumber;
    }
}
