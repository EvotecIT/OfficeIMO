using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private static readonly ConcurrentDictionary<WordprocessingDocumentType, DefaultPackageStyles> DefaultPackageStyleCatalogs = new();

    // Default catalogs with library styles appended, keyed by document type and the appended style IDs in order.
    private static readonly ConcurrentDictionary<(WordprocessingDocumentType Type, string AppendedStyles), DefaultPackageStyles> ExtendedPackageStyleCatalogs = new();

    // The style part copied from the cached default package; while unparsed it still holds that package's styles,
    // followed by the library styles named in _defaultPackageAppendedStyles.
    private StyleDefinitionsPart? _defaultPackageStylesPart;
    private string _defaultPackageAppendedStyles = string.Empty;

    /// <summary>Style IDs and dependencies of a cached default package, and its parsed style catalog to clone.</summary>
    private sealed class DefaultPackageStyles {
        internal DefaultPackageStyles(Styles styles, byte[]? partBytes = null) {
            Styles = styles;
            PartBytes = partBytes;
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

        /// <summary>The style part as the Open XML SDK saves <see cref="Styles"/>; <see langword="null"/> for the
        /// default package's own catalog, which the copied part already holds.</summary>
        internal byte[]? PartBytes { get; }

        internal Dictionary<string, string?[]> Dependencies { get; } = new(StringComparer.Ordinal);
    }

    /// <summary>Returns the default package's style catalog while this document's style part is still that unparsed copy.</summary>
    private DefaultPackageStyles? GetUnparsedDefaultPackageStyles() {
        StyleDefinitionsPart? part = _defaultPackageStylesPart;
        if (part == null || part.IsRootElementLoaded ||
            !ReferenceEquals(_wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart, part)) {
            return null;
        }
        DefaultPackageStyles defaults = DefaultPackageStyleCatalogs.GetOrAdd(_defaultPackageType, ReadDefaultPackageStyles);
        return _defaultPackageAppendedStyles.Length == 0
            ? defaults
            : ExtendedPackageStyleCatalogs.TryGetValue((_defaultPackageType, _defaultPackageAppendedStyles), out DefaultPackageStyles? extended)
                ? extended
                : null;
    }

    /// <summary>
    /// Appends a library style to an unparsed default style part by writing the part the Open XML SDK would save after
    /// appending it to the parsed catalog, so the part stays unparsed and is not serialized again on every save.
    /// </summary>
    /// <param name="styleId">Style that the unparsed catalog does not define.</param>
    /// <param name="dependencies">The appended style's based-on, linked and next styles.</param>
    /// <returns><see langword="false"/> when the part is not an unparsed default catalog or the style is not a
    /// theme-independent library style; the caller then takes the parsed route.</returns>
    private bool TryAppendToUnparsedDefaultPackageStyles(string styleId, out string?[] dependencies) {
        dependencies = Array.Empty<string?>();
        if (GetUnparsedDefaultPackageStyles() is not { } catalog) return false;
        // The modern table style is generated from the document's creation theme.
        if (WordTableStyles.TryGetStyle(styleId, out WordTableStyle tableStyle) && tableStyle == WordTableStyle.OfficeIMOModern) return false;

        var key = (_defaultPackageType, _defaultPackageAppendedStyles + "\n" + styleId);
        if (!ExtendedPackageStyleCatalogs.TryGetValue(key, out DefaultPackageStyles? extended)) {
            Style? definition = FindLibraryStyleDefinition(styleId);
            if (definition == null) return false;
            var styles = (Styles)catalog.Styles.CloneNode(true);
            // Selecting a style for content must preserve both existing defaults and their absence.
            definition.Default = null;
            styles.Append(definition);
            using var serialized = new MemoryStream();
            styles.Save(serialized);
            extended = ExtendedPackageStyleCatalogs.GetOrAdd(key, new DefaultPackageStyles(styles, serialized.ToArray()));
        }

        using (var stream = new MemoryStream(extended.PartBytes!, writable: false)) {
            _defaultPackageStylesPart!.FeedData(stream);
        }
        _defaultPackageAppendedStyles = key.Item2;
        dependencies = extended.Dependencies[styleId];
        return true;
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
