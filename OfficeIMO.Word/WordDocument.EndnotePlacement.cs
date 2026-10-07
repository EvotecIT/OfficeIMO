using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Gets the effective document-wide placement; Word ignores section-level positions.</summary>
    internal EndnotePositionValues DocumentEndnotePosition =>
        _wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<EndnoteDocumentWideProperties>()?.GetFirstChild<EndnotePosition>()?.Val?.Value
        ?? EndnotePositionValues.DocumentEnd;

    /// <summary>Updates placement without changing document or section numbering properties.</summary>
    internal void SetDocumentEndnotePosition(EndnotePositionValues position) {
        _ = Settings;
        var settings = _wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        var properties = settings.GetFirstChild<EndnoteDocumentWideProperties>();
        if (properties == null) {
            properties = new EndnoteDocumentWideProperties();
            settings.AddChild(properties, true);
        }
        properties.RemoveAllChildren<EndnotePosition>();
        properties.AddChild(new EndnotePosition { Val = position }, true);
    }
}
