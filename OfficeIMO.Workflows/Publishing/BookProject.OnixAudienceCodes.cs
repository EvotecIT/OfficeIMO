using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAudienceCodes(IReadOnlyList<BookOnixAudienceCode> codes, CancellationToken token) {
        var result = new List<XElement>();
        var identities = new HashSet<(BookOnixAudienceScheme, string?, string?, string?)>();
        var mainSchemes = new HashSet<BookOnixAudienceScheme>();
        XNamespace ns = OnixNamespace;
        foreach (var code in codes) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(code);
            string scheme = code.Scheme switch {
                BookOnixAudienceScheme.Proprietary => "02", BookOnixAudienceScheme.Btlf => "06",
                BookOnixAudienceScheme.Electre => "07", BookOnixAudienceScheme.Anele => "08",
                BookOnixAudienceScheme.Avi => "09", BookOnixAudienceScheme.Aws => "11",
                BookOnixAudienceScheme.FinnishSchoolLevel => "15", BookOnixAudienceScheme.CbgAgeGuidance => "16",
                BookOnixAudienceScheme.BookData => "17", BookOnixAudienceScheme.AviRevised => "18",
                BookOnixAudienceScheme.JapaneseChildren => "21", BookOnixAudienceScheme.Cefr => "23",
                BookOnixAudienceScheme.IntendedLanguage => "27", BookOnixAudienceScheme.SwedishCurriculum => "29",
                BookOnixAudienceScheme.Isced2011 => "30", _ => throw new ArgumentOutOfRangeException(nameof(code.Scheme))
            };
            var headings = BuildOnixAudienceHeadings(code.Headings, token);
            if (code.Value == null && headings.Count == 0)
                throw new ArgumentException("An audience declaration needs a code value, a heading, or both.", nameof(codes));
            if (code.Value != null) {
                RequireOnixText(code.Value, nameof(code.Value));
                if (code.Value != code.Value.Trim()) throw new ArgumentException("Audience code values cannot have surrounding whitespace.", nameof(codes));
            }
            if (code.Scheme == BookOnixAudienceScheme.Proprietary) {
                RequireOnixText(code.SchemeName!, nameof(code.SchemeName));
                if (code.SchemeName != code.SchemeName!.Trim())
                    throw new ArgumentException("Audience scheme names cannot have surrounding whitespace.", nameof(codes));
            } else if (code.SchemeName != null)
                throw new ArgumentException("Only proprietary audience codes carry a scheme name.", nameof(codes));
            if (code.Value != null && code.Scheme == BookOnixAudienceScheme.Cefr && code.Value is not ("A1" or "A2" or "B1" or "B2" or "C1" or "C2"))
                throw new ArgumentException("CEFR audience codes must be A1, A2, B1, B2, C1 or C2.", nameof(codes));
            if (code.Value != null && code.Scheme == BookOnixAudienceScheme.JapaneseChildren &&
                (code.Value.Length != 2 || code.Value.Any(value => value < '0' || value > '9')))
                throw new ArgumentException("Japanese children's audience codes require two ASCII digits.", nameof(codes));
            if (code.Value != null && code.Scheme == BookOnixAudienceScheme.IntendedLanguage)
                RequireOnixLanguageCode(code.Value, nameof(code.Value));
            // Compare uncoded assertions by their language/text pairs without changing output order.
            string? headingIdentity = code.Value == null ? string.Concat(headings
                .OrderBy(heading => (string?)heading.Attribute("language") ?? string.Empty, StringComparer.Ordinal)
                .Select(heading => heading.ToString(SaveOptions.DisableFormatting))) : null;
            if (!identities.Add((code.Scheme, code.SchemeName, code.Value, headingIdentity)))
                throw new ArgumentException("Audience declarations must be distinct within their named scheme.", nameof(codes));
            if (code.IsMain && !mainSchemes.Add(code.Scheme))
                throw new ArgumentException("At most one main audience is permitted per audience code type, including proprietary schemes.", nameof(codes));
            result.Add(new XElement(ns + "Audience", code.IsMain ? new XElement(ns + "MainAudience") : null,
                new XElement(ns + "AudienceCodeType", scheme),
                code.SchemeName != null ? new XElement(ns + "AudienceCodeTypeName", code.SchemeName) : null,
                code.Value != null ? new XElement(ns + "AudienceCodeValue", code.Value) : null, headings));
        }
        return result;
    }
}
