using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private static readonly ConcurrentDictionary<WordprocessingDocumentType, DefaultPackageStyles> DefaultPackageStyleCatalogs = new();

    // The style part copied from the cached default package; while unparsed it still holds that package's styles.
    private StyleDefinitionsPart? _defaultPackageStylesPart;

    /// <summary>Style IDs and dependencies of a cached default package, and its parsed style catalog to clone.</summary>
    private sealed class DefaultPackageStyles {
        internal DefaultPackageStyles(Styles styles) {
            Styles = styles;
            foreach (Style style in styles.Elements<Style>()) {
                string? id = style.StyleId?.Value;
                if (string.IsNullOrEmpty(id) || Dependencies.ContainsKey(id!)) continue;
                Dependencies.Add(id!, new[] {
                    style.BasedOn?.Val?.Value,
                    style.LinkedStyle?.Val?.Value,
                    style.NextParagraphStyle?.Val?.Value
                });
            }
        }

        internal Styles Styles { get; }

        internal Dictionary<string, string?[]> Dependencies { get; } = new(StringComparer.Ordinal);
    }

    /// <summary>Returns the default package's style catalog while this document's style part is still that unparsed copy.</summary>
    private DefaultPackageStyles? GetUnparsedDefaultPackageStyles() {
        StyleDefinitionsPart? part = _defaultPackageStylesPart;
        if (part == null || part.IsRootElementLoaded ||
            !ReferenceEquals(_wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart, part)) {
            return null;
        }
        return DefaultPackageStyleCatalogs.GetOrAdd(_defaultPackageType, ReadDefaultPackageStyles);
    }

    /// <summary>
    /// Gives an unparsed default style part a copy of the cached catalog, which is cheaper than parsing the same XML
    /// again for every new document that adds a style.
    /// </summary>
    private void LoadDefaultPackageStyles() {
        if (GetUnparsedDefaultPackageStyles() is { } defaults) {
            _defaultPackageStylesPart!.Styles = (Styles)defaults.Styles.CloneNode(true);
        }
    }

    private static DefaultPackageStyles ReadDefaultPackageStyles(WordprocessingDocumentType documentType) {
        using var stream = new MemoryStream(GetDefaultThemedPackageBytes(documentType), writable: false);
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        Styles styles = package.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?? throw new InvalidOperationException("The default Word package has no style catalog.");
        return new DefaultPackageStyles((Styles)styles.CloneNode(true));
    }
}
