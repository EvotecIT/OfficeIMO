using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private static readonly ConcurrentDictionary<WordprocessingDocumentType, A.ColorScheme?> DefaultPackageColorSchemes = new();

    private A.ColorScheme? _retainedColorScheme;
    // The theme part copied from the cached default package; its color scheme is known without parsing it.
    private ThemePart? _defaultPackageThemePart;
    private WordprocessingDocumentType _defaultPackageType;

    /// <summary>
    /// Reads the current native theme, or its retained DOM after disposal, so inspection results can construct wrappers lazily.
    /// </summary>
    internal A.ColorScheme? ReadColorScheme => _wordprocessingDocument == null
        ? _retainedColorScheme
        : _wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    private void RetainColorContextForInspection() {
        ThemePart? themePart = _wordprocessingDocument.MainDocumentPart?.ThemePart;
        // Every theme edit goes through the part's DOM, so an unloaded default theme still has the cached package's
        // colors. Reuse them instead of parsing the theme of every disposed new document.
        if (themePart != null && !themePart.IsRootElementLoaded && ReferenceEquals(themePart, _defaultPackageThemePart)) {
            _retainedColorScheme = DefaultPackageColorSchemes.GetOrAdd(_defaultPackageType, ReadDefaultPackageColorScheme);
            return;
        }
        _retainedColorScheme = ReadColorScheme;
    }

    private static A.ColorScheme? ReadDefaultPackageColorScheme(WordprocessingDocumentType documentType) {
        using var stream = new MemoryStream(GetDefaultThemedPackageBytes(documentType), writable: false);
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme?.CloneNode(true) as A.ColorScheme;
    }
}
