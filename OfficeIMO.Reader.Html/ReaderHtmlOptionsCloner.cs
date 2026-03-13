using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Reader.Html;

internal static class ReaderHtmlOptionsCloner {
    public static ReaderHtmlOptions CloneOrDefault(ReaderHtmlOptions? options) {
        return new ReaderHtmlOptions {
            HtmlToMarkdownOptions = Clone(options?.HtmlToMarkdownOptions) ?? new HtmlToMarkdownOptions()
        };
    }

    public static ReaderHtmlOptions? CloneNullable(ReaderHtmlOptions? options) {
        if (options == null) return null;
        return new ReaderHtmlOptions {
            HtmlToMarkdownOptions = Clone(options.HtmlToMarkdownOptions)
        };
    }

    public static HtmlToMarkdownOptions? Clone(HtmlToMarkdownOptions? options) {
        if (options == null) return null;
        return new HtmlToMarkdownOptions {
            BaseUri = options.BaseUri,
            UseBodyContentsOnly = options.UseBodyContentsOnly,
            RemoveScriptsAndStyles = options.RemoveScriptsAndStyles,
            PreserveUnsupportedBlocks = options.PreserveUnsupportedBlocks,
            PreserveUnsupportedInlineHtml = options.PreserveUnsupportedInlineHtml
        };
    }
}
