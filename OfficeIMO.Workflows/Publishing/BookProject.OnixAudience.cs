using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAudience(BookOnixAudienceMetadata? audience, CancellationToken cancellationToken) {
        if (audience == null) return [];
        ArgumentNullException.ThrowIfNull(audience.Categories);
        ArgumentNullException.ThrowIfNull(audience.Codes);
        ArgumentNullException.ThrowIfNull(audience.AdultRatings);
        ArgumentNullException.ThrowIfNull(audience.AgeRanges);
        ArgumentNullException.ThrowIfNull(audience.GradeRanges);
        ArgumentNullException.ThrowIfNull(audience.Descriptions);
        if (audience.Categories.Count > 13 || audience.Codes.Count > 64 || audience.AdultRatings.Count > 14 || audience.AgeRanges.Count > 3 || audience.GradeRanges.Count > 3 || audience.Descriptions.Count > 16)
            throw new ArgumentException("Audience metadata exceeds its category, code, rating, range or description limit.", nameof(audience));
        if (audience.Categories.Count == 0 && audience.Codes.Count == 0 && audience.AdultRatings.Count == 0 && audience.AgeRanges.Count == 0 && audience.GradeRanges.Count == 0 && audience.Descriptions.Count == 0)
            throw new ArgumentException("Supply at least one audience assertion, or omit Audience.", nameof(audience));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var categories = new HashSet<BookOnixAudienceType>();
        bool hasMain = false;
        foreach (var category in audience.Categories) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(category);
            if (!categories.Add(category.Type) || (category.IsMain && hasMain))
                throw new ArgumentException("Audience categories must be distinct, with at most one main category.", nameof(audience));
            hasMain |= category.IsMain;
            string code = category.Type switch {
                BookOnixAudienceType.GeneralAdult => "01", BookOnixAudienceType.Children => "02",
                BookOnixAudienceType.Teenage => "03", BookOnixAudienceType.PrimaryAndSecondaryEducation => "04",
                BookOnixAudienceType.TertiaryEducation => "05", BookOnixAudienceType.ProfessionalAndScholarly => "06",
                BookOnixAudienceType.EnglishLanguageTeaching => "07", BookOnixAudienceType.AdultEducation => "08",
                BookOnixAudienceType.AdditionalLanguageTeaching => "09", BookOnixAudienceType.PrePrimaryEducation => "11",
                BookOnixAudienceType.PrimaryEducation => "12", BookOnixAudienceType.LowerSecondaryEducation => "13",
                BookOnixAudienceType.UpperSecondaryEducation => "14", _ => throw new ArgumentOutOfRangeException(nameof(category.Type))
            };
            result.Add(new XElement(ns + "Audience", category.IsMain ? new XElement(ns + "MainAudience") : null,
                new XElement(ns + "AudienceCodeType", "01"), new XElement(ns + "AudienceCodeValue", code),
                BuildOnixAudienceHeadings(category.Headings, cancellationToken)));
        }
        result.AddRange(BuildOnixAudienceCodes(audience.Codes, cancellationToken));
        result.AddRange(BuildOnixAdultAudience(audience.AdultRatings, categories.Contains(BookOnixAudienceType.GeneralAdult), cancellationToken));
        var rangeTypes = new HashSet<BookOnixAgeRangeType>();
        foreach (var range in audience.AgeRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(range);
            if (!rangeTypes.Add(range.Type)) throw new ArgumentException("Age range types must be distinct.", nameof(audience));
            result.Add(BuildOnixAgeRange(range));
        }
        if (rangeTypes.Contains(BookOnixAgeRangeType.InterestMonths) && rangeTypes.Contains(BookOnixAgeRangeType.InterestYears))
            throw new ArgumentException("Interest age cannot be specified in both months and years.", nameof(audience));
        var gradeSystems = new HashSet<BookOnixGradeSystem>();
        foreach (var range in audience.GradeRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(range);
            if (!gradeSystems.Add(range.System)) throw new ArgumentException("Grade systems must be distinct.", nameof(audience));
            result.Add(BuildOnixGradeRange(range));
        }
        var languages = new HashSet<string>(StringComparer.Ordinal);
        foreach (var description in audience.Descriptions) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(description);
            RequireOnixText(description.Text, nameof(description.Text));
            RequireOnixTranslationLanguage(description.LanguageCode, audience.Descriptions.Count, nameof(audience.Descriptions));
            if (!languages.Add(description.LanguageCode ?? ""))
                throw new ArgumentException("Audience description languages must be distinct.", nameof(audience));
            var element = new XElement(ns + "AudienceDescription", new XAttribute("textformat", "06"), description.Text);
            if (description.LanguageCode != null) element.Add(new XAttribute("language", description.LanguageCode));
            result.Add(element);
        }
        return result;
    }

    private static XElement BuildOnixAgeRange(BookOnixAgeRange range) {
        if ((range.Minimum == null && range.Maximum == null) || range.Minimum < 0 || range.Maximum < 0 || range.Minimum > range.Maximum)
            throw new ArgumentException("Supply ordered nonnegative age bounds.", nameof(range));
        string qualifier = range.Type switch {
            BookOnixAgeRangeType.InterestMonths => "16", BookOnixAgeRangeType.InterestYears => "17",
            BookOnixAgeRangeType.ReadingYears => "18", _ => throw new ArgumentOutOfRangeException(nameof(range.Type))
        };
        bool exact = range.Minimum != null && range.Minimum == range.Maximum;
        int first = range.Minimum ?? range.Maximum!.Value;
        bool hasSecond = range.Minimum != null && range.Maximum != null && !exact;
        if (range.Type == BookOnixAgeRangeType.InterestMonths && (first > 36 || (hasSecond && range.Maximum > 42)))
            throw new ArgumentException("Interest months permits a first value up to 36 and a second value up to 42.", nameof(range));
        XNamespace ns = OnixNamespace;
        var element = new XElement(ns + "AudienceRange", new XElement(ns + "AudienceRangeQualifier", qualifier),
            new XElement(ns + "AudienceRangePrecision", exact ? "01" : range.Minimum != null ? "03" : "04"),
            new XElement(ns + "AudienceRangeValue", first));
        if (hasSecond) element.Add(new XElement(ns + "AudienceRangePrecision", "04"), new XElement(ns + "AudienceRangeValue", range.Maximum!.Value));
        return element;
    }
}
