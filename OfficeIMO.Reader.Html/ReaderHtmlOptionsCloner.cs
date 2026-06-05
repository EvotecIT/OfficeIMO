using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Reader.Html;

internal static class ReaderHtmlOptionsCloner {
    public static ReaderHtmlOptions CloneOrDefault(ReaderHtmlOptions? options) {
        return options?.Clone() ?? ReaderHtmlOptions.CreateOfficeIMOProfile();
    }

    public static ReaderHtmlOptions? CloneNullable(ReaderHtmlOptions? options) {
        return options?.Clone();
    }
}
