using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class PdfRtfConverter {
    private const int HeadingStyleBaseId = 1;
    private const int BulletListId = 1;
    private const int DecimalListId = 2;

    public static RtfDocument Convert(PdfCore.PdfLogicalDocument source, PdfRtfReadOptions? options) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        PdfRtfReadOptions readOptions = options?.Clone() ?? new PdfRtfReadOptions();
        RtfDocument document = RtfDocument.Create();

        if (readOptions.IncludeMetadata) {
            CopyMetadata(source.Metadata, document.Info);
        }

        if (source.Pages.Count > 0) {
            ApplyFirstPageSetup(source.Pages[0], document);
        }

        if (readOptions.CreateHeadingStyles && readOptions.ImportHeadings) {
            AddHeadingStyles(document);
        }

        for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
            ImportPage(source.Pages[pageIndex], document, readOptions, pageIndex > 0 && readOptions.PreservePageBreaks);
        }

        return document;
    }

    public static RtfDocument Convert(PdfCore.PdfReadDocument source, PdfRtfReadOptions? options) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }

        PdfRtfReadOptions readOptions = options?.Clone() ?? new PdfRtfReadOptions();
        return Convert(PdfCore.PdfLogicalDocument.From(source, readOptions.LayoutOptions), readOptions);
    }

    private static void ImportPage(PdfCore.PdfLogicalPage page, RtfDocument document, PdfRtfReadOptions options, bool pageBreakBeforeFirstParagraph) {
        var consumed = new HashSet<PdfCore.PdfLogicalTextBlock>();
        bool emittedPageContent = false;
        bool pendingPageBreak = pageBreakBeforeFirstParagraph;

        foreach (PdfCore.PdfLogicalTextBlock block in page.TextBlocks) {
            if (consumed.Contains(block)) {
                continue;
            }

            RtfParagraph? paragraph = null;
            if (options.ImportHeadings && TryFindHeading(page, block, out PdfCore.PdfLogicalHeading? heading)) {
                paragraph = AddHeading(document, heading!);
                consumed.Add(block);
            } else if (options.ImportLists && TryFindListItem(page, block, out PdfCore.PdfLogicalListItem? listItem)) {
                paragraph = AddListItem(document, listItem!);
                consumed.Add(block);
            } else if (TryFindParagraph(page, block, out PdfCore.PdfLogicalParagraph? logicalParagraph)) {
                paragraph = AddParagraph(document, logicalParagraph!.Text);
                foreach (PdfCore.PdfLogicalTextBlock line in logicalParagraph.Lines) {
                    consumed.Add(line);
                }
            } else if (!string.IsNullOrWhiteSpace(block.Text)) {
                paragraph = AddParagraph(document, block.Text);
                consumed.Add(block);
            }

            if (paragraph is null) {
                continue;
            }

            if (pendingPageBreak) {
                paragraph.PageBreakBefore = true;
                pendingPageBreak = false;
            }

            emittedPageContent = true;
        }

        if (!emittedPageContent && pendingPageBreak && options.IncludeEmptyPages) {
            document.AddParagraph().PageBreakBefore = true;
        }
    }

    private static RtfParagraph AddHeading(RtfDocument document, PdfCore.PdfLogicalHeading heading) {
        int level = Math.Max(1, Math.Min(3, heading.Level));
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetStyle(HeadingStyleBaseId + level - 1);
        paragraph.SetOutlineLevel(level - 1);
        paragraph.SetParagraphSpacing(afterTwips: 120);
        paragraph.SetPagination(keepWithNext: true);

        RtfRun run = paragraph.AddText(heading.Text);
        run.SetBold();
        if (heading.FontSize > 0) {
            run.SetFontSize(heading.FontSize);
        }

        return paragraph;
    }

    private static RtfParagraph AddListItem(RtfDocument document, PdfCore.PdfLogicalListItem item) {
        RtfListKind kind = IsBulletMarker(item.Marker) ? RtfListKind.Bullet : RtfListKind.Decimal;
        int listId = kind == RtfListKind.Bullet ? BulletListId : DecimalListId;
        int level = Math.Max(0, item.Level - 1);
        RtfParagraph paragraph = AddParagraph(document, item.Text);
        paragraph.SetList(listId, level, kind);
        paragraph.SetListText(GetListMarkerText(item.Marker, kind));
        paragraph.SetIndentation(leftTwips: 720 + (level * 360), firstLineTwips: -360);
        return paragraph;
    }

    private static RtfParagraph AddParagraph(RtfDocument document, string text) {
        RtfParagraph paragraph = document.AddParagraph(text);
        paragraph.SetParagraphSpacing(afterTwips: 120);
        return paragraph;
    }

    private static void CopyMetadata(PdfCore.PdfMetadata source, RtfDocumentInfo target) {
        target.Title = source.Title;
        target.Author = source.Author;
        target.Subject = source.Subject;
        target.Keywords = source.Keywords;
    }

    private static void ApplyFirstPageSetup(PdfCore.PdfLogicalPage page, RtfDocument document) {
        int widthTwips = PointsToTwips(page.Width);
        int heightTwips = PointsToTwips(page.Height);
        if (widthTwips > 0 && heightTwips > 0) {
            document.PageSetup.SetPaperSize(widthTwips, heightTwips);
        }
    }

    private static void AddHeadingStyles(RtfDocument document) {
        AddHeadingStyle(document, 1, 24D);
        AddHeadingStyle(document, 2, 18D);
        AddHeadingStyle(document, 3, 14D);
    }

    private static void AddHeadingStyle(RtfDocument document, int level, double fontSize) {
        RtfStyle style = document.AddStyle(HeadingStyleBaseId + level - 1, "Heading " + level.ToString(System.Globalization.CultureInfo.InvariantCulture));
        style.Bold = true;
        style.FontSize = fontSize;
        style.OutlineLevel = level - 1;
        style.SpaceAfterTwips = 120;
        style.KeepWithNext = true;
    }

    private static bool TryFindHeading(PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalTextBlock block, out PdfCore.PdfLogicalHeading? heading) {
        foreach (PdfCore.PdfLogicalHeading candidate in page.Headings) {
            if (ReferenceEquals(candidate.Line, block)) {
                heading = candidate;
                return true;
            }
        }

        heading = null;
        return false;
    }

    private static bool TryFindListItem(PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalTextBlock block, out PdfCore.PdfLogicalListItem? listItem) {
        foreach (PdfCore.PdfLogicalListItem candidate in page.ListItems) {
            if (ReferenceEquals(candidate.Line, block)) {
                listItem = candidate;
                return true;
            }
        }

        listItem = null;
        return false;
    }

    private static bool TryFindParagraph(PdfCore.PdfLogicalPage page, PdfCore.PdfLogicalTextBlock block, out PdfCore.PdfLogicalParagraph? paragraph) {
        foreach (PdfCore.PdfLogicalParagraph candidate in page.Paragraphs) {
            if (candidate.Lines.Contains(block)) {
                paragraph = candidate;
                return true;
            }
        }

        paragraph = null;
        return false;
    }

    private static string GetListMarkerText(string marker, RtfListKind kind) {
        string normalized = string.IsNullOrWhiteSpace(marker)
            ? kind == RtfListKind.Bullet ? "\u2022" : "1."
            : marker.Trim();

        return normalized + "\t";
    }

    private static bool IsBulletMarker(string marker) {
        string trimmed = marker.Trim();
        if (trimmed.Length == 0) {
            return false;
        }

        return trimmed == "\u2022" ||
            trimmed == "\u25CF" ||
            trimmed == "-" ||
            trimmed == "*" ||
            trimmed == "\u00B7";
    }

    private static int PointsToTwips(double points) {
        if (points <= 0 || double.IsNaN(points) || double.IsInfinity(points)) {
            return 0;
        }

        return (int)Math.Round(points * 20D, MidpointRounding.AwayFromZero);
    }
}
