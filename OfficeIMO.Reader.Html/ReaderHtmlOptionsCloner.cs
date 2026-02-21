using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Reader.Html;

internal static class ReaderHtmlOptionsCloner {
    public static ReaderHtmlOptions CloneOrDefault(ReaderHtmlOptions? options) {
        return new ReaderHtmlOptions {
            HtmlToWordOptions = Clone(options?.HtmlToWordOptions) ?? new HtmlToWordOptions(),
            MarkdownOptions = Clone(options?.MarkdownOptions) ?? new WordToMarkdownOptions()
        };
    }

    public static ReaderHtmlOptions? CloneNullable(ReaderHtmlOptions? options) {
        if (options == null) return null;
        return new ReaderHtmlOptions {
            HtmlToWordOptions = Clone(options.HtmlToWordOptions),
            MarkdownOptions = Clone(options.MarkdownOptions)
        };
    }

    public static HtmlToWordOptions? Clone(HtmlToWordOptions? options) {
        if (options == null) return null;
        var clone = new HtmlToWordOptions {
            FontFamily = options.FontFamily,
            QuotePrefix = options.QuotePrefix,
            QuoteSuffix = options.QuoteSuffix,
            DefaultPageSize = options.DefaultPageSize,
            DefaultOrientation = options.DefaultOrientation,
            IncludeListStyles = options.IncludeListStyles,
            ContinueNumbering = options.ContinueNumbering,
            SupportsHeadingNumbering = options.SupportsHeadingNumbering,
            BasePath = options.BasePath,
            NoteReferenceType = options.NoteReferenceType,
            LinkNoteUrls = options.LinkNoteUrls,
            ImageProcessing = options.ImageProcessing,
            HttpClient = options.HttpClient,
            ResourceTimeout = options.ResourceTimeout,
            RenderPreAsTable = options.RenderPreAsTable,
            TableCaptionPosition = options.TableCaptionPosition,
            SectionTagHandling = options.SectionTagHandling
        };

        foreach (var item in options.ClassStyles) {
            clone.ClassStyles[item.Key] = item.Value;
        }

        foreach (var stylesheetPath in options.StylesheetPaths) {
            clone.StylesheetPaths.Add(stylesheetPath);
        }

        foreach (var stylesheet in options.StylesheetContents) {
            clone.StylesheetContents.Add(stylesheet);
        }

        return clone;
    }

    public static WordToMarkdownOptions? Clone(WordToMarkdownOptions? options) {
        if (options == null) return null;
        return new WordToMarkdownOptions {
            FontFamily = options.FontFamily,
            EnableUnderline = options.EnableUnderline,
            EnableHighlight = options.EnableHighlight,
            ImageExportMode = options.ImageExportMode,
            ImageDirectory = options.ImageDirectory
        };
    }
}
