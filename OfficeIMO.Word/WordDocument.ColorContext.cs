using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private A.ColorScheme? _retainedColorScheme;

    /// <summary>
    /// Reads the current native theme, or its retained DOM after disposal, so inspection results can construct wrappers lazily.
    /// </summary>
    internal A.ColorScheme? ReadColorScheme => _wordprocessingDocument == null
        ? _retainedColorScheme
        : _wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    private void RetainColorContextForInspection() => _retainedColorScheme = ReadColorScheme;
}
