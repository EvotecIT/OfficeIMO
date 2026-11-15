using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// The package parts an OfficeIMO document carries besides its body and header/footer stories: styles, settings,
/// web settings, font table, theme, footnote and endnote separators, and the core and extended properties.
/// </summary>
/// <remarks>
/// The Open XML SDK <c>Equivalent</c> lanes write these parts so that both libraries produce the same package; the
/// <c>Minimal</c> lanes write only the main document (and its stories), which is less work than OfficeIMO does.
/// The parts are captured once, outside measurement, from OfficeIMO's own output for the same workload, and are fed
/// to the SDK as serialized XML: the cheapest way an SDK caller can produce them.
/// </remarks>
internal sealed class WordEquivalentPackageParts {
    private readonly List<(string RelationshipId, string ContentType, byte[] Data)> _mainDocumentParts;
    private readonly byte[]? _coreProperties;
    private readonly byte[]? _extendedProperties;

    private WordEquivalentPackageParts(
        List<(string RelationshipId, string ContentType, byte[] Data)> mainDocumentParts,
        byte[]? coreProperties,
        byte[]? extendedProperties) {
        _mainDocumentParts = mainDocumentParts;
        _coreProperties = coreProperties;
        _extendedProperties = extendedProperties;
    }

    /// <summary>Captures the supporting parts of an OfficeIMO package.</summary>
    internal static WordEquivalentPackageParts Capture(byte[] package) {
        using var stream = new MemoryStream(package, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        MainDocumentPart mainPart = document.MainDocumentPart
            ?? throw new InvalidDataException("The reference package has no main document part.");
        var parts = new List<(string, string, byte[])>();
        foreach (IdPartPair pair in mainPart.Parts) {
            if (pair.OpenXmlPart is HeaderPart or FooterPart) continue;
            parts.Add((pair.RelationshipId, pair.OpenXmlPart.ContentType, ReadPart(pair.OpenXmlPart)));
        }
        return new WordEquivalentPackageParts(
            parts,
            document.CoreFilePropertiesPart is { } core ? ReadPart(core) : null,
            document.ExtendedFilePropertiesPart is { } extended ? ReadPart(extended) : null);
    }

    /// <summary>Adds the captured parts to a new SDK package that already has its main document part.</summary>
    internal void AddTo(WordprocessingDocument document, MainDocumentPart mainPart) {
        foreach ((string relationshipId, string contentType, byte[] data) in _mainDocumentParts) {
            Feed(AddMainDocumentChild(mainPart, contentType, relationshipId), data);
        }
        if (_coreProperties != null) Feed(document.AddCoreFilePropertiesPart(), _coreProperties);
        if (_extendedProperties != null) Feed(document.AddExtendedFilePropertiesPart(), _extendedProperties);
    }

    /// <summary>Fails when two packages differ in their part names or content types.</summary>
    internal static void EnsureSamePartSet(string name, byte[] expected, byte[] actual) {
        string[] expectedParts = DescribeParts(expected);
        string[] actualParts = DescribeParts(actual);
        if (!expectedParts.SequenceEqual(actualParts, StringComparer.Ordinal)) {
            throw new InvalidDataException(
                name + " writes parts [" + string.Join(", ", actualParts) + "]; OfficeIMO writes [" +
                string.Join(", ", expectedParts) + "].");
        }
    }

    private static string[] DescribeParts(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return archive.Entries
            .Where(entry => !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.Ordinal))
            .Select(entry => entry.FullName)
            .Select(name => name.StartsWith("word/header", StringComparison.Ordinal) ? "word/header" :
                name.StartsWith("word/footer", StringComparison.Ordinal) ? "word/footer" : name)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();
    }

    private static OpenXmlPart AddMainDocumentChild(MainDocumentPart mainPart, string contentType, string relationshipId) =>
        contentType switch {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml" => mainPart.AddNewPart<StyleDefinitionsPart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" => mainPart.AddNewPart<DocumentSettingsPart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml" => mainPart.AddNewPart<WebSettingsPart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" => mainPart.AddNewPart<FontTablePart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.theme+xml" => mainPart.AddNewPart<ThemePart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml" => mainPart.AddNewPart<FootnotesPart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" => mainPart.AddNewPart<EndnotesPart>(relationshipId),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" => mainPart.AddNewPart<NumberingDefinitionsPart>(relationshipId),
            _ => throw new InvalidDataException("The equivalent SDK package has no mapping for main document part content type " + contentType + ".")
        };

    private static byte[] ReadPart(OpenXmlPart part) {
        using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static void Feed(OpenXmlPart part, byte[] data) {
        using var input = new MemoryStream(data, writable: false);
        part.FeedData(input);
    }
}
