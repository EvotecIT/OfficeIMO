using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
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
