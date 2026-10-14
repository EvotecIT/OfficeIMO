namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkKeynoteArchiveBuilder {
    private void AddDefaultStyles() {
        AddTextStyles(DefaultCharacterId, DefaultParagraphId, DefaultFrameId, "Arial", 32,
            new IWorkColor(0, 0, 0, 255));
        var list = P().Message(1, Style(ListId, "None")).UInt(10, 3);
        for (int index = 0; index < 9; index++) list.UInt(11, 0).Float(12, 0).Float(13, 0);
        Add(ListId, 2023, list);
        Add(TextPresetId, 2050, P().String(1, "officeimo-text").Reference(2, DefaultParagraphId).Reference(3, ListId));
        AddBackground(BlankBackgroundId, new IWorkColor(255, 255, 255, 255));
    }

    private void AddBackground(ulong id, IWorkColor color) =>
        Add(id, 9, P().Message(1, Style(id, "OfficeIMO Slide")).Message(11, P().Message(1, P().Message(1, Color(color)))).UInt(10, 1));

    private void AddTheme() {
        var drawingPresets = P().Reference(6, DefaultFrameId).Reference(5, DefaultFrameId)
            .Reference(4, DefaultFrameId).Reference(9, DefaultFrameId);
        if (_hasImageStyle) drawingPresets.Reference(7, ImageStyleId);
        var textPresets = P().Reference(2, TextPresetId).Reference(7, DefaultParagraphId)
            .Reference(6, DefaultCharacterId).Reference(1, ListId);
        var theme = P().Reference(4, StylesheetId).String(3, "OfficeIMO")
            .Message(100, drawingPresets).Message(110, textPresets);
        AddTableThemePresets(theme);
        IWorkNativeThemeWriter.AddColorPresets(theme, _limits.MaximumPackageBytes, _cancellationToken);
        Add(ThemeId, 10, P().Message(1, theme).String(3, IWorkCreationIdentity.Create(_modelHash, "theme").Text)
            .Reference(2, BlankNodeId).Reference(5, BlankNodeId).Reference(6, BlankNodeId).Bool(7, false));
    }

}
