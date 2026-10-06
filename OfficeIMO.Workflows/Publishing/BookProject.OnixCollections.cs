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
            var title = BuildOnixCollectionTitle(collection, cancellationToken, out var titleLevels);
            if (collection.SourceName != null || collection.Type == BookOnixCollectionType.Ascribed)
                RequireOnixText(collection.SourceName!, nameof(collection.SourceName));
            string type = collection.Type switch {
                BookOnixCollectionType.Publisher => "10", BookOnixCollectionType.Editorial => "11",
                BookOnixCollectionType.Ascribed => "20", _ => throw new ArgumentOutOfRangeException(nameof(collection.Type))
            };
            var element = new XElement(ns + "Collection", new XElement(ns + "CollectionType", type));
            if (collection.Frequency is { } frequency)
                element.Add(new XElement(ns + "CollectionFrequency", OnixCollectionFrequencyCode(frequency)));
            if (collection.SourceName != null) element.Add(new XElement(ns + "SourceName", collection.SourceName));
            element.Add(BuildOnixCollectionIdentifiers(collection.Identifiers, titleLevels, cancellationToken));
            element.Add(BuildOnixCollectionSequences(collection.Sequences, cancellationToken));
            element.Add(title);
            element.Add(BuildOnixContributors(collection.Contributors, collection.NoContributors,
                requireDeclaration: false, cancellationToken));
            result.Add(element);
        }
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixCollectionIdentifiers(IReadOnlyList<BookOnixCollectionIdentifier> identifiers,
        IReadOnlyCollection<BookOnixCollectionLevel> titleLevels, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(identifiers);
        if (identifiers.Count > 16) throw new ArgumentException("At most 16 collection identifiers are supported.", nameof(identifiers));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var schemes = new HashSet<(BookOnixCollectionIdentifierType Type, string? Name, BookOnixCollectionLevel? Level)>();
        foreach (var identifier in identifiers) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(identifier);
            RequireOnixText(identifier.Value, nameof(identifier.Value));
            bool proprietary = identifier.Type == BookOnixCollectionIdentifierType.Proprietary;
            RequireOnixNamedScheme(proprietary, identifier.SchemeName, nameof(identifier.SchemeName));
            string? level = identifier.Level is { } selectedLevel ? OnixCollectionLevelCode(selectedLevel) : null;
            if (identifier.Level is { } requiredLevel && !titleLevels.Contains(requiredLevel))
                throw new ArgumentException("An identifier's collection level must occur in the collection title.", nameof(identifiers));
            if (schemes.Any(scheme => scheme.Type == identifier.Type && scheme.Name == identifier.SchemeName &&
                (scheme.Level == identifier.Level || scheme.Level == null || identifier.Level == null)))
                throw new ArgumentException("Collection identifier schemes must be distinct within each level; scoped and unscoped values cannot overlap.", nameof(identifiers));
            schemes.Add((identifier.Type, identifier.SchemeName, identifier.Level));
            (string type, string value) = identifier.Type switch {
                BookOnixCollectionIdentifierType.Proprietary => ("01", identifier.Value),
                BookOnixCollectionIdentifierType.Issn => ("02", OfficeIssn.Normalize(identifier.Value)),
                BookOnixCollectionIdentifierType.Isbn13 => ("15", OfficeIsbn.Normalize(identifier.Value, true)),
                BookOnixCollectionIdentifierType.GermanNationalBibliography => ("03", identifier.Value),
                BookOnixCollectionIdentifierType.GermanBooksInPrint => ("04", identifier.Value),
                BookOnixCollectionIdentifierType.Electre => ("05", identifier.Value),
                BookOnixCollectionIdentifierType.Doi => ("06", RequireOnixCollectionDoi(identifier.Value)),
                BookOnixCollectionIdentifierType.Urn => ("22", RequireOnixCollectionUrn(identifier.Value)),
                BookOnixCollectionIdentifierType.JapaneseMagazine => ("27", RequireOnixMagazineIdentifier(identifier.Value)),
                BookOnixCollectionIdentifierType.BnfControlNumber => ("29", identifier.Value),
                BookOnixCollectionIdentifierType.Ark => ("35", RequireOnixCollectionArk(identifier.Value)),
                BookOnixCollectionIdentifierType.IssnL => ("38", OfficeIssn.Normalize(identifier.Value)),
                _ => throw new ArgumentOutOfRangeException(nameof(identifier.Type))
            };
            result.Add(new XElement(ns + "CollectionIdentifier", level != null ? new XElement(ns + "CollectionElementLevel", level) : null,
                new XElement(ns + "CollectionIDType", type),
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
