using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static void RenderFrontMatter(PdfCore.PdfDocument pdf, FrontMatterBlock frontMatter, MarkdownDoc document, MarkdownPdfSaveOptions options, MarkdownPdfStyle visualTheme) {
        if (frontMatter.Entries.Count == 0) {
            return;
        }

        switch (options.FrontMatterRenderMode) {
            case MarkdownPdfFrontMatterRenderMode.Hidden:
                return;
            case MarkdownPdfFrontMatterRenderMode.DocumentHeader:
                if (RenderFrontMatterDocumentHeader(pdf, frontMatter, document, visualTheme)) {
                    return;
                }

                break;
            case MarkdownPdfFrontMatterRenderMode.Table:
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options.FrontMatterRenderMode), options.FrontMatterRenderMode, "Unsupported Markdown PDF front matter render mode.");
        }

        RenderFrontMatterTable(pdf, frontMatter, visualTheme);
    }

    private static bool RenderFrontMatterDocumentHeader(PdfCore.PdfDocument pdf, FrontMatterBlock frontMatter, MarkdownDoc document, MarkdownPdfStyle visualTheme) {
        string? title = GetFrontMatterMetadata(frontMatter, "title");
        if (title == null) {
            return false;
        }

        string? anchor = FindMatchingFirstHeadingAnchor(document, title);
        if (!string.IsNullOrWhiteSpace(anchor)) {
            pdf.Bookmark(anchor!);
        }

        pdf.H1(title, PdfCore.PdfAlign.Left, visualTheme.DocumentHeaderTitleColorSnapshot, style: new PdfCore.PdfHeadingStyle {
            FontSize = visualTheme.DocumentHeaderTitleFontSizeSnapshot,
            LineHeight = 1.12,
            SpacingBefore = 0,
            SpacingAfter = 3,
            Color = visualTheme.DocumentHeaderTitleColorSnapshot,
            KeepWithNext = true
        });

        string? subtitle = GetFrontMatterMetadata(frontMatter, "subtitle")
            ?? GetFrontMatterMetadata(frontMatter, "description")
            ?? GetFrontMatterMetadata(frontMatter, "summary")
            ?? GetFrontMatterMetadata(frontMatter, "subject");
        if (subtitle != null) {
            pdf.Paragraph(builder => builder
                    .FontSize(visualTheme.DocumentHeaderSubtitleFontSizeSnapshot)
                    .Color(visualTheme.DocumentHeaderSubtitleColorSnapshot)
                    .Text(subtitle),
                defaultColor: visualTheme.DocumentHeaderSubtitleColorSnapshot,
                style: new PdfCore.PdfParagraphStyle {
                    LineHeight = 1.25,
                    SpacingAfter = 4,
                    KeepWithNext = true
                });
        }

        string? metadataLine = BuildFrontMatterMetadataLine(frontMatter);
        if (metadataLine != null) {
            pdf.Paragraph(builder => builder
                    .FontSize(visualTheme.DocumentHeaderMetadataFontSizeSnapshot)
                    .Color(visualTheme.DocumentHeaderMetadataColorSnapshot)
                    .Text(metadataLine),
                defaultColor: visualTheme.DocumentHeaderMetadataColorSnapshot,
                style: new PdfCore.PdfParagraphStyle {
                    LineHeight = 1.2,
                    SpacingAfter = 6,
                    KeepWithNext = true
                });
        }

        pdf.HR(style: new PdfCore.PdfHorizontalRuleStyle {
            Color = visualTheme.DocumentHeaderRuleColorSnapshot,
            Thickness = 0.8,
            SpacingBefore = 2,
            SpacingAfter = 12,
            KeepWithNext = false
        });
        return true;
    }

    private static void RenderFrontMatterTable(PdfCore.PdfDocument pdf, FrontMatterBlock frontMatter, MarkdownPdfStyle visualTheme) {
        var rows = new List<PdfCore.PdfKeyValueRow>();
        for (int i = 0; i < frontMatter.Entries.Count; i++) {
            rows.Add(PdfCore.PdfKeyValueRow.Text(frontMatter.Entries[i].Key, ConvertMetadataValue(frontMatter.Entries[i].Value)));
        }

        PdfCore.PdfTableStyle style = visualTheme.FrontMatterTableStyleSnapshot;
        pdf.KeyValueTable(rows, style: style, includeHeader: true);
    }

    private static string? BuildFrontMatterMetadataLine(FrontMatterBlock frontMatter) {
        var parts = new List<string>();
        AddMetadataPart(parts, GetFrontMatterMetadata(frontMatter, "author"));
        AddMetadataPart(parts, GetFrontMatterMetadata(frontMatter, "date") ?? GetFrontMatterMetadata(frontMatter, "published") ?? GetFrontMatterMetadata(frontMatter, "updated"));
        string? tags = GetFrontMatterMetadata(frontMatter, "tags") ?? GetFrontMatterMetadata(frontMatter, "keywords");
        if (tags != null) {
            AddMetadataPart(parts, "Tags: " + tags);
        }

        return parts.Count == 0 ? null : string.Join(" | ", parts);
    }

    private static void AddMetadataPart(List<string> parts, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            parts.Add(value!.Trim());
        }
    }
}
