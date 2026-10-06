using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement[] BuildOnixSalesRestrictions(IReadOnlyList<BookOnixSalesRestriction> restrictions,
        ref int budget, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(restrictions);
        if (restrictions.Count > 32) throw new ArgumentException("At most 32 sales restrictions are supported per territory or market.", nameof(restrictions));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>(); var prior = new List<BookOnixSalesRestriction>();
        foreach (var restriction in restrictions) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(restriction);
            if (!Enum.IsDefined(restriction.Kind)) throw new ArgumentOutOfRangeException(nameof(restriction.Kind));
            ArgumentNullException.ThrowIfNull(restriction.Outlets); ArgumentNullException.ThrowIfNull(restriction.Notes);
            if (restriction.Outlets.Count > 16 || restriction.Notes.Count > 16)
                throw new ArgumentException("At most 16 outlets and 16 translated notes are supported per restriction.", nameof(restrictions));
            if (restriction.ValidFrom > restriction.ValidUntil) throw new ArgumentException("Restriction dates cannot run backwards.", nameof(restrictions));
            if (restriction.Kind == BookOnixSalesRestrictionKind.Unspecified && restriction.Notes.Count == 0)
                throw new ArgumentException("Unspecified restrictions require an explanatory note.", nameof(restrictions));
            if (restriction.Kind is BookOnixSalesRestrictionKind.RetailerExclusiveOrOwnBrand or BookOnixSalesRestrictionKind.RetailerExclusive or
                BookOnixSalesRestrictionKind.RetailerOwnBrand or BookOnixSalesRestrictionKind.RetailerException or
                BookOnixSalesRestrictionKind.SelectedSubscriptionServices or BookOnixSalesRestrictionKind.SubscriptionServiceExclusive && restriction.Outlets.Count == 0)
                throw new ArgumentException("This restriction requires named or identified outlets.", nameof(restrictions));
            if (restriction.Kind == BookOnixSalesRestrictionKind.NoRestrictions && restriction.Outlets.Count != 0)
                throw new ArgumentException("NoRestrictions cannot identify a restricted set of outlets.", nameof(restrictions));
            foreach (var previous in prior) {
                bool overlap = (previous.ValidFrom ?? DateOnly.MinValue) <= (restriction.ValidUntil ?? DateOnly.MaxValue) &&
                    (restriction.ValidFrom ?? DateOnly.MinValue) <= (previous.ValidUntil ?? DateOnly.MaxValue);
                if (overlap && OpposedOnixRestrictions(previous.Kind, restriction.Kind))
                    throw new ArgumentException("Opposing restrictions cannot apply during overlapping date ranges.", nameof(restrictions));
            }
            prior.Add(restriction);
            var element = new XElement(ns + "SalesRestriction", new XElement(ns + "SalesRestrictionType", ((int)restriction.Kind).ToString("D2", CultureInfo.InvariantCulture)));
            foreach (var outlet in restriction.Outlets) {
                token.ThrowIfCancellationRequested();
                element.Add(BuildOnixSalesOutlet(outlet, ref budget));
            }
            var languages = new HashSet<string>(StringComparer.Ordinal);
            foreach (var note in restriction.Notes) {
                token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(note);
                RequireOnixRestrictionText(note.Text, 300, ref budget);
                RequireOnixTranslationLanguage(note.LanguageCode, restriction.Notes.Count, nameof(restriction.Notes));
                if (!languages.Add(note.LanguageCode ?? string.Empty)) throw new ArgumentException("Restriction note languages must be distinct.", nameof(restrictions));
                element.Add(new XElement(ns + "SalesRestrictionNote", new XAttribute("textformat", "06"),
                    note.LanguageCode != null ? new XAttribute("language", note.LanguageCode) : null, note.Text));
            }
            if (restriction.ValidFrom is { } from) element.Add(new XElement(ns + "StartDate", new XAttribute("dateformat", "00"), from.ToString("yyyyMMdd", CultureInfo.InvariantCulture)));
            if (restriction.ValidUntil is { } until) element.Add(new XElement(ns + "EndDate", new XAttribute("dateformat", "00"), until.ToString("yyyyMMdd", CultureInfo.InvariantCulture)));
            result.Add(element);
        }
        return result.ToArray();
    }

    private static XElement BuildOnixSalesOutlet(BookOnixSalesOutlet outlet, ref int budget) {
        ArgumentNullException.ThrowIfNull(outlet); ArgumentNullException.ThrowIfNull(outlet.Identifiers);
        if (outlet.Identifiers.Count > 8 || outlet.Name == null && outlet.Identifiers.Count == 0)
            throw new ArgumentException("An outlet requires a name or identifier, and supports at most eight identifiers.", nameof(outlet));
        if (outlet.Name == null && outlet.NameLanguageCode != null) throw new ArgumentException("Name language requires an outlet name.", nameof(outlet));
        XNamespace ns = OnixNamespace;
        var result = new XElement(ns + "SalesOutlet");
        var schemes = new HashSet<(BookOnixSalesOutletScheme, string?)>();
        foreach (var identifier in outlet.Identifiers) {
            ArgumentNullException.ThrowIfNull(identifier);
            if (!Enum.IsDefined(identifier.Scheme)) throw new ArgumentOutOfRangeException(nameof(identifier.Scheme));
            RequireOnixRestrictionText(identifier.Value, 100, ref budget);
            if (identifier.Scheme == BookOnixSalesOutletScheme.Proprietary) RequireOnixRestrictionText(identifier.SchemeName!, 100, ref budget);
            else if (identifier.SchemeName != null) throw new ArgumentException("Only proprietary outlet identifiers have a scheme name.", nameof(outlet));
            if (!schemes.Add((identifier.Scheme, identifier.SchemeName))) throw new ArgumentException("Outlet identifier schemes must be distinct.", nameof(outlet));
            bool valid = identifier.Scheme switch {
                BookOnixSalesOutletScheme.Gln => identifier.Value.Length == 13 && identifier.Value.All(c => c is >= '0' and <= '9'),
                BookOnixSalesOutletScheme.San => identifier.Value.Length == 7 && identifier.Value.All(c => c is >= '0' and <= '9'),
                BookOnixSalesOutletScheme.Onix => identifier.Value.Length == 3 && identifier.Value.All(c => c is >= 'A' and <= 'Z' or >= '0' and <= '9'),
                _ => true
            };
            if (!valid) throw new ArgumentException("Outlet identifier does not match its scheme's lexical format.", nameof(outlet));
            result.Add(new XElement(ns + "SalesOutletIdentifier", new XElement(ns + "SalesOutletIDType", ((int)identifier.Scheme).ToString("D2", CultureInfo.InvariantCulture)),
                identifier.SchemeName != null ? new XElement(ns + "IDTypeName", identifier.SchemeName) : null, new XElement(ns + "IDValue", identifier.Value)));
        }
        if (outlet.Name != null) {
            RequireOnixRestrictionText(outlet.Name, 200, ref budget);
            RequireOnixTranslationLanguage(outlet.NameLanguageCode, 1, nameof(outlet.NameLanguageCode));
            result.Add(new XElement(ns + "SalesOutletName", outlet.NameLanguageCode != null ? new XAttribute("language", outlet.NameLanguageCode) : null, outlet.Name));
        }
        return result;
    }

    private static void RequireOnixRestrictionText(string value, int maximum, ref int budget) {
        RequireOnixText(value, nameof(value));
        if (value.Length > maximum) throw new ArgumentException("Sales restriction text exceeds its field limit.", nameof(value));
        budget -= value.Length;
        if (budget < 0) throw new ArgumentException("Combined sales restriction text exceeds 524288 UTF-16 code units.", nameof(value));
    }

    private static bool OpposedOnixRestrictions(BookOnixSalesRestrictionKind first, BookOnixSalesRestrictionKind second) {
        if (first == second) return false;
        if (first == BookOnixSalesRestrictionKind.NoRestrictions || second == BookOnixSalesRestrictionKind.NoRestrictions) return true;
        int low = Math.Min((int)first, (int)second), high = Math.Max((int)first, (int)second);
        return (low, high) is (6, 9) or (7, 16) or (12, 13) or (14, 15) or (22, 23);
    }
}
