using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixAlternativeTitles(IReadOnlyList<BookOnixAlternativeTitle> titles, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(titles);
        if (titles.Count > 32) throw new ArgumentException("Supply at most 32 alternative ONIX titles.", nameof(titles));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        foreach (var title in titles) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(title);
            string type = title.Type switch {
                BookOnixAlternativeTitleType.OriginalLanguage => "03", BookOnixAlternativeTitleType.Abbreviated => "05",
                BookOnixAlternativeTitleType.OtherLanguage => "06", BookOnixAlternativeTitleType.Former => "08",
                BookOnixAlternativeTitleType.Distributor => "10", BookOnixAlternativeTitleType.Cover => "11",
                BookOnixAlternativeTitleType.BackCover => "12", BookOnixAlternativeTitleType.Expanded => "13",
                BookOnixAlternativeTitleType.Alternative => "14", BookOnixAlternativeTitleType.Spine => "15",
                BookOnixAlternativeTitleType.TranslatedFrom => "16", _ => throw new ArgumentOutOfRangeException(nameof(title.Type))
            };
            RequireOnixText(title.Title, nameof(title.Title));
            if (title.LanguageCode != null) RequireOnixLanguageCode(title.LanguageCode, nameof(title.LanguageCode));
            if (title.Subtitle != null) RequireOnixText(title.Subtitle, nameof(title.Subtitle));
            var element = new XElement(ns + "TitleElement", new XElement(ns + "TitleElementLevel", "01"),
                BuildOnixTitleText(title.Title, title.TitleSorting, title.LanguageCode));
            if (title.Subtitle != null) element.Add(new XElement(ns + "Subtitle",
                title.LanguageCode != null ? new XAttribute("language", title.LanguageCode) : null, title.Subtitle));
            result.Add(new XElement(ns + "TitleDetail", new XElement(ns + "TitleType", type), element));
        }
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixTitleText(string? title, BookOnixTitleSorting? sorting, string? language = null) {
        XNamespace ns = OnixNamespace;
        if (title == null) {
            if (sorting != null) throw new ArgumentException("Title sorting requires title text; it cannot apply to a part designation alone.", nameof(sorting));
            return [];
        }
        RequireOnixText(title, nameof(title));
        if (sorting == null) return [Text("TitleText", title)];
        if (sorting.Prefix == null) return [new XElement(ns + "NoPrefix"), Text("TitleWithoutPrefix", title)];
        RequireOnixText(sorting.Prefix, nameof(sorting.Prefix));
        if (!title.StartsWith(sorting.Prefix, StringComparison.Ordinal))
            throw new ArgumentException("The sorting prefix must match the exact beginning of the selected title.", nameof(sorting));
        string remainder = title.Substring(sorting.Prefix.Length);
        RequireOnixText(remainder, nameof(sorting));
        return [Text("TitlePrefix", sorting.Prefix), Text("TitleWithoutPrefix", remainder)];

        XElement Text(string name, string value) => new(ns + name,
            language != null ? new XAttribute("language", language) : null, value);
    }
}
