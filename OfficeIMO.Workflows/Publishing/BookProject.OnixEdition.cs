using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixEdition(BookOnixEdition? edition, CancellationToken cancellationToken) {
        if (edition == null) return [];
        ArgumentNullException.ThrowIfNull(edition.Types);
        ArgumentNullException.ThrowIfNull(edition.Statements);
        if (edition.Types.Count > 8 || edition.Statements.Count > 16)
            throw new ArgumentException("Edition metadata permits at most eight types and 16 statements.", nameof(edition));
        bool hasDetails = edition.Types.Count != 0 || edition.Number != null || edition.VersionNumber != null || edition.Statements.Count != 0;
        if (edition.NoEdition == hasDetails)
            throw new ArgumentException("Supply edition details or explicitly assert NoEdition.", nameof(edition));
        XNamespace ns = OnixNamespace;
        if (edition.NoEdition) return [new XElement(ns + "NoEdition")];
        if (edition.Number is <= 0 || (edition.VersionNumber != null && edition.Number == null))
            throw new ArgumentException("Edition numbers must be positive; a version requires a numbered edition.", nameof(edition));
        var result = new List<XElement>();
        var types = new HashSet<BookOnixEditionType>();
        foreach (var type in edition.Types) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!types.Add(type)) throw new ArgumentException("Edition types must be distinct.", nameof(edition));
            string code = type switch {
                BookOnixEditionType.Abridged => "ABR", BookOnixEditionType.Unabridged => "UBR",
                BookOnixEditionType.Annotated => "ANN", BookOnixEditionType.Revised => "REV",
                BookOnixEditionType.Enlarged => "ENL", BookOnixEditionType.Illustrated => "ILL",
                BookOnixEditionType.Critical => "CRI", BookOnixEditionType.New => "NED",
                _ => throw new ArgumentOutOfRangeException(nameof(edition.Types))
            };
            result.Add(new XElement(ns + "EditionType", code));
        }
        if (types.Contains(BookOnixEditionType.Abridged) && types.Contains(BookOnixEditionType.Unabridged))
            throw new ArgumentException("An edition cannot be both abridged and unabridged.", nameof(edition));
        if (types.Contains(BookOnixEditionType.New) && (types.Count != 1 || edition.Number != null))
            throw new ArgumentException("New edition is used only when no more specific type or numbering applies.", nameof(edition));
        if (edition.Number != null) result.Add(new XElement(ns + "EditionNumber", edition.Number.Value));
        if (edition.VersionNumber != null) {
            RequireOnixText(edition.VersionNumber, nameof(edition.VersionNumber));
            result.Add(new XElement(ns + "EditionVersionNumber", edition.VersionNumber));
        }
        var languages = new HashSet<string>(StringComparer.Ordinal);
        foreach (var statement in edition.Statements) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(statement);
            RequireOnixText(statement.Text, nameof(statement.Text));
            if (statement.LanguageCode != null) RequireOnixLanguageCode(statement.LanguageCode, nameof(statement.LanguageCode));
            if (!languages.Add(statement.LanguageCode ?? ""))
                throw new ArgumentException("Edition statement languages must be distinct.", nameof(edition));
            var element = new XElement(ns + "EditionStatement", new XAttribute("textformat", "06"), statement.Text);
            if (statement.LanguageCode != null) element.Add(new XAttribute("language", statement.LanguageCode));
            result.Add(element);
        }
        return result;
    }
}
