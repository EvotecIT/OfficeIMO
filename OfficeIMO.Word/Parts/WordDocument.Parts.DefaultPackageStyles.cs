using System.Collections.Concurrent;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private static readonly ConcurrentDictionary<WordprocessingDocumentType, DefaultPackageStyles> DefaultPackageStyleCatalogs = new();

    // Default catalogs with one library style appended, keyed by document type and style. Only the first style a
    // document adds is served from here and the entries keep no parsed catalog, so the cache stays small and bounded.
    private const int MaxExtendedPackageStyleCatalogs = 32;
    private static readonly ConcurrentDictionary<(WordprocessingDocumentType Type, string StyleId), DefaultPackageStyles> ExtendedPackageStyleCatalogs = new();
    // Admissions reserved or stored, so concurrent misses cannot take the cache past its limit.
    private static int _extendedPackageStyleCatalogSlots;

    // The style part copied from the cached default package; while unparsed it still holds that package's styles,
    // followed by _defaultPackageAppendedStyle when one was appended.
    private StyleDefinitionsPart? _defaultPackageStylesPart;
    private string? _defaultPackageAppendedStyle;

    /// <summary>Style IDs and dependencies of a cached default package, and its parsed style catalog to clone.</summary>
    private sealed class DefaultPackageStyles {
        internal DefaultPackageStyles(Styles styles, byte[]? partBytes = null) {
            // An extended catalog is served as part bytes; a document that needs its DOM parses the part.
            Styles = partBytes == null ? styles : null;
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

        /// <summary>The parsed default catalog to clone; <see langword="null"/> for an extended catalog.</summary>
        internal Styles? Styles { get; }

        /// <summary>The style part as the Open XML SDK saves the extended catalog; <see langword="null"/> for the
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
        if (_defaultPackageAppendedStyle == null) {
            return DefaultPackageStyleCatalogs.GetOrAdd(_defaultPackageType, ReadDefaultPackageStyles);
        }
        return ExtendedPackageStyleCatalogs.TryGetValue((_defaultPackageType, _defaultPackageAppendedStyle), out DefaultPackageStyles? extended)
            ? extended
            : null;
    }

    /// <summary>
    /// Appends a library style to an unparsed default style part by writing the part the Open XML SDK would save after
    /// appending it to the parsed catalog, so the part stays unparsed and is not serialized again on every save.
    /// </summary>
    /// <param name="styleId">Style that the unparsed catalog does not define.</param>
    /// <param name="dependencies">The appended style's based-on, linked and next styles.</param>
    /// <returns><see langword="false"/> when the part is not the unparsed default catalog, already has a style
    /// appended, the style is not a theme-independent library style or the cache is full; the caller then takes the
    /// parsed route.</returns>
    private bool TryAppendToUnparsedDefaultPackageStyles(string styleId, out string?[] dependencies) {
        dependencies = Array.Empty<string?>();
        if (GetUnparsedDefaultPackageStyles() is not { Styles: { } defaultStyles }) return false;
        // The modern table style is generated from the document's creation theme.
        if (WordTableStyles.TryGetStyle(styleId, out WordTableStyle tableStyle) && tableStyle == WordTableStyle.OfficeIMOModern) return false;

        var key = (_defaultPackageType, styleId);
        if (!ExtendedPackageStyleCatalogs.TryGetValue(key, out DefaultPackageStyles? extended)) {
            if (Interlocked.Increment(ref _extendedPackageStyleCatalogSlots) > MaxExtendedPackageStyleCatalogs) {
                Interlocked.Decrement(ref _extendedPackageStyleCatalogSlots);
                return false;
            }
            DefaultPackageStyles? created = null;
            try {
                Style? definition = FindLibraryStyleDefinition(styleId);
                if (definition == null) return false;
                var styles = (Styles)defaultStyles.CloneNode(true);
                // Selecting a style for content must preserve both existing defaults and their absence.
                definition.Default = null;
                styles.Append(definition);
                using var serialized = new MemoryStream();
                styles.Save(serialized);
                created = new DefaultPackageStyles(styles, serialized.ToArray());
                extended = ExtendedPackageStyleCatalogs.GetOrAdd(key, created);
            } finally {
                // Release the slot unless this call stored the entry.
                if (created == null || !ReferenceEquals(extended, created)) Interlocked.Decrement(ref _extendedPackageStyleCatalogSlots);
            }
        }

        using (var stream = new MemoryStream(extended.PartBytes!, writable: false)) {
            _defaultPackageStylesPart!.FeedData(stream);
        }
        _defaultPackageAppendedStyle = styleId;
        dependencies = extended.Dependencies[styleId];
        return true;
    }

    /// <summary>
    /// Gives an unparsed default style part a copy of the cached catalog, which is cheaper than parsing the same XML
    /// again for every new document that adds a style. A part extended with a library style is parsed when needed.
    /// </summary>
    private void LoadDefaultPackageStyles() {
        if (GetUnparsedDefaultPackageStyles() is { Styles: { } styles }) {
            _defaultPackageStylesPart!.Styles = (Styles)styles.CloneNode(true);
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
