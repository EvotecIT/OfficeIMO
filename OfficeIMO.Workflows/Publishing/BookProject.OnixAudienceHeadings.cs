using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAudienceHeadings(IReadOnlyList<BookOnixAudienceHeading> headings, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(headings);
        if (headings.Count > 16) throw new ArgumentException("Supply at most 16 audience heading translations.", nameof(headings));
        var result = new List<XElement>();
        var languages = new HashSet<string>(StringComparer.Ordinal);
        XNamespace ns = OnixNamespace;
        foreach (var heading in headings) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(heading);
            RequireOnixText(heading.Text, nameof(heading.Text));
            if (headings.Count > 1 && heading.LanguageCode == null)
                throw new ArgumentException("Repeated audience headings require a language on every translation.", nameof(headings));
            if (heading.LanguageCode != null) RequireOnixLanguageCode(heading.LanguageCode, nameof(heading.LanguageCode));
            if (!languages.Add(heading.LanguageCode ?? string.Empty))
                throw new ArgumentException("Audience heading languages must be distinct.", nameof(headings));
            var element = new XElement(ns + "AudienceHeadingText", heading.Text);
            if (heading.LanguageCode != null) element.SetAttributeValue("language", heading.LanguageCode);
            result.Add(element);
        }
        return result;
    }
}
