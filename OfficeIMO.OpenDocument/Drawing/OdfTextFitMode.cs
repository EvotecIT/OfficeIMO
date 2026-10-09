namespace OfficeIMO.OpenDocument;

/// <summary>Native fitting of text within an OpenDocument shape's saved frame.</summary>
public enum OdfTextFitMode {
    /// <summary>Uses the declared font sizes without fitting text to the frame.</summary>
    None,
    /// <summary>Stretches text to the frame; shared drawing projection does not reproduce this native mode.</summary>
    Stretch,
    /// <summary>Reduces text size to fit a fixed frame, without enlarging its saved geometry.</summary>
    ShrinkToFit
}
