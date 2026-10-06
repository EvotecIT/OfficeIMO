using OfficeIMO.Core.Internal;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixCollections(BookOnixExportOptions options, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(options.Collections);
        if (options.Collections.Count > 32 || (options.NoCollection && options.Collections.Count != 0))
            throw new ArgumentException("Supply at most 32 collections, or explicitly assert NoCollection.", nameof(options));
        XNamespace ns = OnixNamespace;
        if (options.NoCollection) return [new XElement(ns + "NoCollection")];
        var result = new List<XElement>();
        foreach (var collection in options.Collections) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(collection);
            RequireOnixText(collection.Title, nameof(collection.Title));
            if (collection.Subtitle != null) RequireOnixText(collection.Subtitle, nameof(collection.Subtitle));
            if (collection.LanguageCode != null) RequireOnixLanguageCode(collection.LanguageCode, nameof(collection.LanguageCode));
            if (collection.SourceName != null || collection.Type == BookOnixCollectionType.Ascribed)
                RequireOnixText(collection.SourceName!, nameof(collection.SourceName));
            string type = collection.Type switch {
                BookOnixCollectionType.Publisher => "10", BookOnixCollectionType.Editorial => "11",
                BookOnixCollectionType.Ascribed => "20", _ => throw new ArgumentOutOfRangeException(nameof(collection.Type))
            };
            var element = new XElement(ns + "Collection", new XElement(ns + "CollectionType", type));
            if (collection.SourceName != null) element.Add(new XElement(ns + "SourceName", collection.SourceName));
            element.Add(BuildOnixCollectionIdentifiers(collection.Identifiers, cancellationToken));
            element.Add(BuildOnixCollectionSequences(collection.Sequences, cancellationToken));
            var title = new XElement(ns + "TitleElement", new XElement(ns + "TitleElementLevel", "02"),
                new XElement(ns + "TitleText", collection.Title));
            if (collection.Subtitle != null) title.Add(new XElement(ns + "Subtitle", collection.Subtitle));
            if (collection.LanguageCode != null) {
                title.Element(ns + "TitleText")!.Add(new XAttribute("language", collection.LanguageCode));
                title.Element(ns + "Subtitle")?.Add(new XAttribute("language", collection.LanguageCode));
            }
            element.Add(new XElement(ns + "TitleDetail", new XElement(ns + "TitleType", "01"), title));
            element.Add(BuildOnixContributors(collection.Contributors, collection.NoContributors,
                requireDeclaration: false, cancellationToken));
            result.Add(element);
        }
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixCollectionIdentifiers(IReadOnlyList<BookOnixCollectionIdentifier> identifiers,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(identifiers);
        if (identifiers.Count > 16) throw new ArgumentException("At most 16 collection identifiers are supported.", nameof(identifiers));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var schemes = new HashSet<(BookOnixCollectionIdentifierType, string?)>();
        foreach (var identifier in identifiers) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(identifier);
            RequireOnixText(identifier.Value, nameof(identifier.Value));
            bool proprietary = identifier.Type == BookOnixCollectionIdentifierType.Proprietary;
            RequireOnixNamedScheme(proprietary, identifier.SchemeName, nameof(identifier.SchemeName));
            if (!schemes.Add((identifier.Type, identifier.SchemeName)))
                throw new ArgumentException("Collection identifier schemes must be distinct.", nameof(identifiers));
            (string type, string value) = identifier.Type switch {
                BookOnixCollectionIdentifierType.Proprietary => ("01", identifier.Value),
                BookOnixCollectionIdentifierType.Issn => ("02", OfficeIssn.Normalize(identifier.Value)),
                BookOnixCollectionIdentifierType.Isbn13 => ("15", OfficeIsbn.Normalize(identifier.Value, true)),
                _ => throw new ArgumentOutOfRangeException(nameof(identifier.Type))
            };
            result.Add(new XElement(ns + "CollectionIdentifier", new XElement(ns + "CollectionIDType", type),
                proprietary ? new XElement(ns + "IDTypeName", identifier.SchemeName) : null, new XElement(ns + "IDValue", value)));
        }
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixCollectionSequences(IReadOnlyList<BookOnixCollectionSequence> sequences,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(sequences);
        if (sequences.Count > 16) throw new ArgumentException("At most 16 collection sequences are supported.", nameof(sequences));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var kinds = new HashSet<(BookOnixCollectionSequenceType, string?)>();
        foreach (var sequence in sequences) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(sequence);
            RequireOnixText(sequence.Number, nameof(sequence.Number));
            if (sequence.Number.Split('.').Any(part => part != "-" && (part.Length == 0 || part.Any(c => c < '0' || c > '9'))))
                throw new ArgumentException("Collection positions must contain dot-separated ASCII integers or hyphens.", nameof(sequence.Number));
            bool proprietary = sequence.Type == BookOnixCollectionSequenceType.Proprietary;
            RequireOnixNamedScheme(proprietary, sequence.Name, nameof(sequence.Name));
            if (!kinds.Add((sequence.Type, sequence.Name)))
                throw new ArgumentException("Collection sequence kinds must be distinct.", nameof(sequences));
            string type = sequence.Type switch {
                BookOnixCollectionSequenceType.Proprietary => "01", BookOnixCollectionSequenceType.Title => "02",
                BookOnixCollectionSequenceType.Publication => "03", BookOnixCollectionSequenceType.Narrative => "04",
                BookOnixCollectionSequenceType.OriginalPublication => "05", BookOnixCollectionSequenceType.SuggestedReading => "06",
                BookOnixCollectionSequenceType.SuggestedDisplay => "07", _ => throw new ArgumentOutOfRangeException(nameof(sequence.Type))
            };
            result.Add(new XElement(ns + "CollectionSequence", new XElement(ns + "CollectionSequenceType", type),
                proprietary ? new XElement(ns + "CollectionSequenceTypeName", sequence.Name) : null,
                new XElement(ns + "CollectionSequenceNumber", sequence.Number)));
        }
        return result;
    }

    private static void RequireOnixNamedScheme(bool proprietary, string? name, string parameterName) {
        if (proprietary) RequireOnixText(name!, parameterName);
        else if (name != null) throw new ArgumentException("Standard schemes do not accept a proprietary name.", parameterName);
    }
}
