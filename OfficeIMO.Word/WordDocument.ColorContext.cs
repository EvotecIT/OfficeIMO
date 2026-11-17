using System.Collections.Concurrent;
using System.Xml;
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
        // An unloaded theme is not parsed into a DOM just to keep its color scheme; only that element is read.
        if (themePart != null && !themePart.IsRootElementLoaded && TryReadUnloadedColorScheme(themePart, out A.ColorScheme? scheme)) {
            _retainedColorScheme = scheme;
            return;
        }
        _retainedColorScheme = ReadColorScheme;
    }

    /// <summary>
    /// Reads <c>a:theme/a:themeElements/a:clrScheme</c> (first of each, as the typed accessors resolve it) from the
    /// part stream. The rest of the part is still read so malformed XML falls back to the DOM and its exception.
    /// </summary>
    private static bool TryReadUnloadedColorScheme(ThemePart themePart, out A.ColorScheme? scheme) {
        const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        scheme = null;
        try {
            using Stream stream = themePart.GetStream(FileMode.Open, FileAccess.Read);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = false
            });
            if (reader.MoveToContent() != XmlNodeType.Element || reader.NamespaceURI != DrawingNamespace || reader.LocalName != "theme") return false;
            string? outerXml = null;
            bool seenThemeElements = false;
            if (!reader.IsEmptyElement) {
                int rootDepth = reader.Depth;
                reader.Read();
                while (!reader.EOF && reader.Depth > rootDepth) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        reader.Read();
                        continue;
                    }
                    if (reader.Depth == rootDepth + 1 && !seenThemeElements &&
                        reader.NamespaceURI == DrawingNamespace && reader.LocalName == "themeElements") {
                        seenThemeElements = true;
                        if (reader.IsEmptyElement) {
                            reader.Read();
                            continue;
                        }
                        int elementsDepth = reader.Depth;
                        reader.Read();
                        while (!reader.EOF && reader.Depth > elementsDepth) {
                            if (reader.NodeType != XmlNodeType.Element) {
                                reader.Read();
                            } else if (outerXml == null && reader.Depth == elementsDepth + 1 &&
                                reader.NamespaceURI == DrawingNamespace && reader.LocalName == "clrScheme") {
                                outerXml = reader.ReadOuterXml();
                            } else {
                                reader.Skip();
                            }
                        }
                        continue;
                    }
                    reader.Skip();
                }
            }
            while (reader.Read()) {
            }
            scheme = outerXml == null ? null : new A.ColorScheme(outerXml);
            return true;
        } catch (XmlException) {
            return false;
        }
    }

    private static A.ColorScheme? ReadDefaultPackageColorScheme(WordprocessingDocumentType documentType) {
        using var stream = new MemoryStream(GetDefaultThemedPackageBytes(documentType), writable: false);
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme?.CloneNode(true) as A.ColorScheme;
    }
}
