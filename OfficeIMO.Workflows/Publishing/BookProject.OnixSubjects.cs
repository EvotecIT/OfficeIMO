using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixSubjects(IReadOnlyList<BookOnixSubject> subjects, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(subjects);
        if (subjects.Count > 64) throw new ArgumentException("Supply at most 64 ONIX subject declarations.", nameof(subjects));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var mainSchemes = new HashSet<(BookOnixSubjectScheme, string?)>();
        foreach (BookOnixSubject subject in subjects) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(subject);
            string scheme = subject.Scheme switch {
                BookOnixSubjectScheme.Dewey => "01", BookOnixSubjectScheme.LibraryOfCongressClassification => "03",
                BookOnixSubjectScheme.LibraryOfCongressHeading => "04", BookOnixSubjectScheme.Bisac => "10",
                BookOnixSubjectScheme.Keywords => "20", BookOnixSubjectScheme.Proprietary => "24",
                BookOnixSubjectScheme.Thema => "93", BookOnixSubjectScheme.ThemaGeographical => "94",
                BookOnixSubjectScheme.ThemaLanguage => "95", BookOnixSubjectScheme.ThemaTimePeriod => "96",
                BookOnixSubjectScheme.ThemaEducationalPurpose => "97", BookOnixSubjectScheme.ThemaInterest => "98",
                BookOnixSubjectScheme.ThemaStyle => "99", _ => throw new ArgumentOutOfRangeException(nameof(subject.Scheme))
            };
            if (subject.Scheme == BookOnixSubjectScheme.Proprietary) RequireOnixText(subject.SchemeName!, nameof(subject.SchemeName));
            else if (subject.SchemeName != null) throw new ArgumentException("Only proprietary subject schemes carry a scheme name.", nameof(subjects));
            if (subject.SchemeVersion != null) RequireOnixText(subject.SchemeVersion, nameof(subject.SchemeVersion));
            if (subject.Code != null) RequireOnixText(subject.Code, nameof(subject.Code));
            ArgumentNullException.ThrowIfNull(subject.Headings);
            if (subject.Headings.Count > 16 || subject.Code == null && subject.Headings.Count == 0 ||
                subject.Scheme == BookOnixSubjectScheme.Keywords && subject.Code != null)
                throw new ArgumentException("Subjects need a code or 1–16 headings; keywords use headings only.", nameof(subjects));
            bool qualifier = subject.Scheme is BookOnixSubjectScheme.ThemaGeographical or BookOnixSubjectScheme.ThemaLanguage or
                BookOnixSubjectScheme.ThemaTimePeriod or BookOnixSubjectScheme.ThemaEducationalPurpose or
                BookOnixSubjectScheme.ThemaInterest or BookOnixSubjectScheme.ThemaStyle;
            if (subject.IsMain && (qualifier || subject.Scheme == BookOnixSubjectScheme.Keywords ||
                !mainSchemes.Add((subject.Scheme, subject.SchemeName))))
                throw new ArgumentException("Only one main subject per scheme is allowed; keywords and qualifiers cannot be main subjects.", nameof(subjects));
            var element = new XElement(ns + "Subject");
            if (subject.IsMain) element.Add(new XElement(ns + "MainSubject"));
            element.Add(new XElement(ns + "SubjectSchemeIdentifier", scheme));
            if (subject.SchemeName != null) element.Add(new XElement(ns + "SubjectSchemeName", subject.SchemeName));
            if (subject.SchemeVersion != null) element.Add(new XElement(ns + "SubjectSchemeVersion", subject.SchemeVersion));
            if (subject.Code != null) element.Add(new XElement(ns + "SubjectCode", subject.Code));
            var languages = new HashSet<string>(StringComparer.Ordinal);
            foreach (BookOnixSubjectHeading heading in subject.Headings) {
                token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(heading);
                RequireOnixText(heading.Text, nameof(heading.Text));
                if (heading.LanguageCode != null) RequireOnixLanguageCode(heading.LanguageCode, nameof(heading.LanguageCode));
                if (!languages.Add(heading.LanguageCode ?? string.Empty))
                    throw new ArgumentException("Subject headings cannot repeat a language, including unspecified language.", nameof(subjects));
                var value = new XElement(ns + "SubjectHeadingText", heading.Text);
                if (heading.LanguageCode != null) value.SetAttributeValue("language", heading.LanguageCode);
                element.Add(value);
            }
            result.Add(element);
        }
        return result;
    }

    private static void RequireOnixLanguageCode(string value, string name) {
        RequireOnixText(value, name);
        if (value.Length != 3 || value.Any(c => c < 'a' || c > 'z'))
            throw new ArgumentException("Supply a three-letter ONIX list 74 language code.", name);
    }
}
