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

    private IWorkProtoWriter Style(ulong id, string name) {
        string identifier = "officeimo-style-" + id.ToString(System.Globalization.CultureInfo.InvariantCulture);
        _styles.Add((id, identifier));
        return P().String(1, name).String(2, identifier).Reference(5, StylesheetId);
    }

    private void AddTextStyles(ulong characterId, ulong paragraphId, ulong frameId, string fontName,
        float fontSize, IWorkColor color) {
        // Modern Keynote renders tsd_fill; font_color remains its legacy color mirror.
        var character = P().Float(3, fontSize).String(5, fontName).Message(7, Color(color))
            .Message(46, P().Message(1, Color(color)));
        Add(characterId, 2021, P().Message(1, Style(characterId, "OfficeIMO Text")).Message(11, character).UInt(10, 3));
        var paragraph = P().UInt(1, 0).Reference(40, ListId).Message(13, P()).Message(25, P()).Float(4, 36);
        Add(paragraphId, 2022, P().Message(1, Style(paragraphId, "OfficeIMO Paragraph"))
            .Message(11, character).Message(12, paragraph).UInt(10, 8));
        var shapeProperties = P().Message(1, P()).Float(3, 1).Message(2, P().Message(6, P().UInt(1, 2)));
        var shape = P().Message(1, Style(frameId, "OfficeIMO Text Frame")).Message(11, shapeProperties).UInt(10, 3);
        var textProperties = P().Bool(1, false).UInt(2, 0).Message(4, P().Message(1, P().UInt(1, 1)))
            .Message(6, P().Float(1, 0).Float(2, 0).Float(3, 0).Float(4, 0)).UInt(7, 0).Bool(11, false).Reference(10, paragraphId);
        Add(frameId, 2025, P().Message(1, shape).Message(11, textProperties).UInt(10, 7));
    }

    private void AddBackground(ulong id, IWorkColor color) =>
        Add(id, 9, P().Message(1, Style(id, "OfficeIMO Slide")).Message(11, P().Message(1, P().Message(1, Color(color)))).UInt(10, 1));

    private void AddTheme() {
        var drawingPresets = P().Reference(6, DefaultFrameId).Reference(5, DefaultFrameId)
            .Reference(4, DefaultFrameId).Reference(9, DefaultFrameId);
        var textPresets = P().Reference(2, TextPresetId).Reference(7, DefaultParagraphId)
            .Reference(6, DefaultCharacterId).Reference(1, ListId);
        var theme = P().Reference(4, StylesheetId).String(3, "OfficeIMO")
            .Message(100, drawingPresets).Message(110, textPresets);
        // Keynote's color-preset UI addresses a 30-slot palette. Generate our own palette, including the no-fill slot.
        foreach (float grey in new[] { 1f, .8f, .6f, .3f, 0f }) theme.Message(10, PaletteColor(grey, grey, grey, 1));
        theme.Message(10, PaletteColor(0, 0, 0, 0));
        foreach (var rgb in new[] { (0f, .5f, 1f), (0f, .8f, .5f), (.5f, .8f, 0f),
                     (1f, .8f, 0f), (1f, .2f, .2f), (.6f, .2f, 1f) }) {
            foreach (float scale in new[] { 1f, .8f, .6f, .4f }) {
                theme.Message(10, PaletteColor(rgb.Item1 * scale, rgb.Item2 * scale, rgb.Item3 * scale, 1));
            }
        }
        Add(ThemeId, 10, P().Message(1, theme).String(3, IWorkKeynoteIdentity.Create(_modelHash, "theme").Text)
            .Reference(2, BlankNodeId).Reference(5, BlankNodeId).Reference(6, BlankNodeId).Bool(7, false));
    }

    private IWorkProtoWriter PaletteColor(float r, float g, float b, float alpha) =>
        P().UInt(1, 1).Float(3, r).Float(4, g).Float(5, b).Float(6, alpha).UInt(12, 1);
}
