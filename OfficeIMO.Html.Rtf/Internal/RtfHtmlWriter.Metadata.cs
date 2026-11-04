using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlWriter {
    private static void AppendDocumentMetadata(StringBuilder builder, RtfDocument document, string newline) {
        AppendMeta(builder, newline, "author", document.Properties.Author);
        AppendMeta(builder, newline, "keywords", document.Properties.Keywords);
        AppendMeta(builder, newline, "description", document.Properties.Subject);
        AppendMeta(builder, newline, "officeimo-rtf-title", document.Properties.Title);
        AppendMeta(builder, newline, "officeimo-rtf-subject", document.Properties.Subject);
        AppendMeta(builder, newline, "officeimo-rtf-author", document.Properties.Author);
        AppendMeta(builder, newline, "officeimo-rtf-manager", document.Properties.Manager);
        AppendMeta(builder, newline, "officeimo-rtf-company", document.Properties.Company);
        AppendMeta(builder, newline, "officeimo-rtf-operator", document.Properties.Operator);
        AppendMeta(builder, newline, "officeimo-rtf-category", document.Properties.Category);
        AppendMeta(builder, newline, "officeimo-rtf-keywords", document.Properties.Keywords);
        AppendMeta(builder, newline, "officeimo-rtf-comments", document.Properties.Comments);
        AppendMeta(builder, newline, "officeimo-rtf-hyperlink-base", document.Properties.HyperlinkBase);
        AppendMeta(builder, newline, "officeimo-rtf-created", FormatTimestamp(document.Properties.Created));
        AppendMeta(builder, newline, "officeimo-rtf-revised", FormatTimestamp(document.Properties.Revised));
        AppendMeta(builder, newline, "officeimo-rtf-printed", FormatTimestamp(document.Properties.Printed));
        AppendMeta(builder, newline, "officeimo-rtf-backed-up", FormatTimestamp(document.Properties.BackedUp));
        AppendMeta(builder, newline, "officeimo-rtf-editing-minutes", document.Properties.EditingMinutes);
        AppendMeta(builder, newline, "officeimo-rtf-pages", document.Properties.NumberOfPages);
        AppendMeta(builder, newline, "officeimo-rtf-words", document.Properties.NumberOfWords);
        AppendMeta(builder, newline, "officeimo-rtf-characters", document.Properties.NumberOfCharacters);
        AppendMeta(builder, newline, "officeimo-rtf-characters-with-spaces", document.Properties.NumberOfCharactersWithSpaces);
        AppendMeta(builder, newline, "officeimo-rtf-internal-version", document.Properties.InternalVersion);
    }

    private static void AppendHeaderFooterMetadata(StringBuilder builder, RtfDocument document, RtfHtmlExportOptions options, string newline) {
        foreach (RtfHeaderFooter headerFooter in document.HeaderFooters) {
            string? content = EncodeHeaderFooterContent(headerFooter, options, document);
            if (content == null) {
                continue;
            }

            builder.Append(newline);
            builder.Append("<meta name=\"officeimo-rtf-header-footer\" content=\"");
            builder.Append(EncodeAttribute(content));
            builder.Append("\" data-officeimo-rtf-kind=\"");
            builder.Append(FormatHeaderFooterKind(headerFooter.Kind));
            for (int index = 0; index < document.Sections.Count; index++) {
                if (!document.Sections[index].HeaderFooters.Contains(headerFooter)) continue;
                builder.Append("\" data-officeimo-rtf-section-index=\"");
                builder.Append(index.ToString(CultureInfo.InvariantCulture));
                break;
            }
            builder.Append("\">");
        }
    }

    private static string? EncodeHeaderFooterContent(RtfHeaderFooter headerFooter, RtfHtmlExportOptions options, RtfDocument document) {
        if (headerFooter.Paragraphs.Count == 0) {
            return string.Empty;
        }

        string newline = options.GetNewLine();
        var builder = new StringBuilder();
        for (int index = 0; index < headerFooter.Paragraphs.Count; index++) {
            if (index > 0) {
                builder.Append(newline);
            }

            AppendParagraph(builder, headerFooter.Paragraphs[index], options, document);
        }

        return EncodeString(builder.ToString());
    }

    private static void AppendMeta(StringBuilder builder, string newline, string name, int? value) {
        if (!value.HasValue) {
            return;
        }

        AppendMeta(builder, newline, name, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static void AppendMeta(StringBuilder builder, string newline, string name, string? value) {
        if (string.IsNullOrEmpty(value)) {
            return;
        }

        builder.Append(newline);
        builder.Append("<meta name=\"");
        builder.Append(EncodeAttribute(name));
        builder.Append("\" content=\"");
        builder.Append(EncodeAttribute(value!));
        builder.Append("\">");
    }

    private static string? FormatTimestamp(DateTime? value) {
        return value?.ToString("O", CultureInfo.InvariantCulture);
    }

    private static string FormatHeaderFooterKind(RtfHeaderFooterKind kind) {
        switch (kind) {
            case RtfHeaderFooterKind.LeftHeader:
                return "left-header";
            case RtfHeaderFooterKind.RightHeader:
                return "right-header";
            case RtfHeaderFooterKind.FirstHeader:
                return "first-header";
            case RtfHeaderFooterKind.Footer:
                return "footer";
            case RtfHeaderFooterKind.LeftFooter:
                return "left-footer";
            case RtfHeaderFooterKind.RightFooter:
                return "right-footer";
            case RtfHeaderFooterKind.FirstFooter:
                return "first-footer";
            default:
                return "header";
        }
    }
}
