using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfConverter {
    internal static PdfCore.PdfDocument Convert(RtfDocument document, RtfPdfSaveOptions? options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        RtfPdfSaveOptions normalized = (options ?? new RtfPdfSaveOptions()).Normalize();
        PdfCore.PdfOptions pdfOptions = normalized.PdfOptions ?? new PdfCore.PdfOptions();
        ApplyPageSetup(document.PageSetup, pdfOptions);

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        ApplyMetadata(document, pdf, normalized);

        foreach (IRtfBlock block in document.Blocks) {
            RenderBlock(document, block, pdf, normalized);
        }

        return pdf;
    }

    private static void RenderBlock(RtfDocument document, IRtfBlock block, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        switch (block) {
            case RtfParagraph paragraph:
                RenderParagraph(document, paragraph, pdf, options);
                break;
            case RtfTable table when options.IncludeTables:
                RenderTable(document, table, pdf, options);
                break;
            case RtfImage image:
                RenderImage(image, pdf, options);
                break;
            case RtfObject rtfObject:
                RenderPlainTextBlock(rtfObject.ToPlainText(), pdf);
                break;
            case RtfShape shape:
                RenderPlainTextBlock(shape.ToPlainText(), pdf);
                break;
        }
    }

    private static void RenderParagraph(RtfDocument document, RtfParagraph paragraph, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        if (paragraph.PageBreakBefore) {
            pdf.PageBreak();
        }

        PdfCore.PdfAlign align = RtfPdfMapping.ToPdfAlign(paragraph.Alignment);
        List<PdfCore.TextRun> pendingRuns = new List<PdfCore.TextRun>();
        bool emitted = false;

        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, pendingRuns, options);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Page:
                    FlushParagraph(pdf, pendingRuns, align);
                    emitted = true;
                    pdf.PageBreak();
                    break;
                case RtfBreak:
                    pendingRuns.Add(PdfCore.TextRun.LineBreak());
                    break;
                case RtfField field:
                    AppendParagraphRuns(document, field.Result, pendingRuns, options);
                    break;
                case RtfImage image:
                    FlushParagraph(pdf, pendingRuns, align);
                    emitted = true;
                    RenderImage(image, pdf, options);
                    break;
                case RtfObject rtfObject:
                    AppendPlainText(rtfObject.ToPlainText(), pendingRuns);
                    break;
                case RtfShape shape:
                    AppendPlainText(shape.ToPlainText(), pendingRuns);
                    break;
                case RtfBookmarkMarker marker when marker.Kind == RtfBookmarkMarkerKind.Start:
                    FlushParagraph(pdf, pendingRuns, align);
                    emitted = true;
                    pdf.Bookmark(marker.Name);
                    break;
            }
        }

        if (pendingRuns.Count > 0 || !emitted) {
            FlushParagraph(pdf, pendingRuns, align);
        }
    }

    private static void RenderTable(RtfDocument document, RtfTable table, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        List<PdfCore.PdfTableCell[]> rows = new List<PdfCore.PdfTableCell[]>();
        foreach (RtfTableRow row in table.Rows) {
            List<PdfCore.PdfTableCell> cells = new List<PdfCore.PdfTableCell>();
            foreach (RtfTableCell cell in row.Cells) {
                if (cell.HorizontalMerge == RtfTableCellMerge.Continue || cell.VerticalMerge == RtfTableCellMerge.Continue) {
                    continue;
                }

                List<PdfCore.TextRun> runs = BuildCellRuns(document, cell, options);
                List<PdfCore.PdfTableCellImage> images = BuildCellImages(cell, options);
                if (images.Count > 0) {
                    cells.Add(PdfCore.PdfTableCell.WithImages(runs, images));
                } else {
                    cells.Add(PdfCore.PdfTableCell.RichTextCell(runs));
                }
            }

            if (cells.Count > 0) {
                rows.Add(cells.ToArray());
            }
        }

        if (rows.Count > 0) {
            pdf.Table(rows);
        }
    }

    private static void RenderImage(RtfImage image, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        if (!options.IncludeImages || !IsPdfSupportedImage(image) || image.Data.Length == 0) {
            return;
        }

        pdf.Image(image.Data, GetImageWidth(image, options), GetImageHeight(image, options), image.Description);
    }

    private static void RenderPlainTextBlock(string text, PdfCore.PdfDocument pdf) {
        if (!string.IsNullOrEmpty(text)) {
            pdf.Paragraph(paragraph => paragraph.Text(text));
        }
    }

    private static void FlushParagraph(PdfCore.PdfDocument pdf, List<PdfCore.TextRun> runs, PdfCore.PdfAlign align) {
        List<PdfCore.TextRun> snapshot = runs.Count == 0
            ? new List<PdfCore.TextRun> { PdfCore.TextRun.Normal(string.Empty) }
            : new List<PdfCore.TextRun>(runs);
        runs.Clear();
        pdf.Paragraph(paragraph => paragraph.Runs(snapshot), align);
    }

    private static void AppendParagraphRuns(RtfDocument document, RtfParagraph paragraph, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options) {
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, runs, options);
                    break;
                case RtfBreak:
                    runs.Add(PdfCore.TextRun.LineBreak());
                    break;
                case RtfField field:
                    AppendParagraphRuns(document, field.Result, runs, options);
                    break;
                case RtfObject rtfObject:
                    AppendPlainText(rtfObject.ToPlainText(), runs);
                    break;
                case RtfShape shape:
                    AppendPlainText(shape.ToPlainText(), runs);
                    break;
            }
        }
    }

    private static void AppendRun(RtfDocument document, RtfRun run, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options) {
        if (run.Hidden && !options.IncludeHiddenText) {
            return;
        }

        string text = run.Text ?? string.Empty;
        if (text.Length == 0) {
            return;
        }

        PdfCore.PdfColor? foreground = RtfPdfMapping.ToPdfColor(document, run.ForegroundColorIndex);
        PdfCore.PdfColor? background = RtfPdfMapping.ToPdfColor(document, run.HighlightColorIndex)
            ?? RtfPdfMapping.ToPdfColor(document, run.CharacterBackgroundColorIndex);
        PdfCore.PdfStandardFont? font = RtfPdfMapping.ToPdfFont(document, run.FontId, run.Bold, run.Italic);

        runs.Add(new PdfCore.TextRun(
            text,
            bold: run.Bold,
            underline: run.Underline,
            color: foreground,
            italic: run.Italic,
            strike: run.Strike || run.DoubleStrike,
            fontSize: run.FontSize,
            font: font,
            linkUri: run.Hyperlink?.ToString(),
            baseline: RtfPdfMapping.ToPdfBaseline(run.VerticalPosition),
            backgroundColor: background));
    }

    private static void AppendPlainText(string text, List<PdfCore.TextRun> runs) {
        if (!string.IsNullOrEmpty(text)) {
            runs.Add(PdfCore.TextRun.Normal(text));
        }
    }

    private static List<PdfCore.TextRun> BuildCellRuns(RtfDocument document, RtfTableCell cell, RtfPdfSaveOptions options) {
        List<PdfCore.TextRun> runs = new List<PdfCore.TextRun>();
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            if (i > 0) {
                runs.Add(PdfCore.TextRun.LineBreak());
            }

            AppendParagraphRuns(document, cell.Paragraphs[i], runs, options);
        }

        if (runs.Count == 0) {
            runs.Add(PdfCore.TextRun.Normal(string.Empty));
        }

        return runs;
    }

    private static List<PdfCore.PdfTableCellImage> BuildCellImages(RtfTableCell cell, RtfPdfSaveOptions options) {
        List<PdfCore.PdfTableCellImage> images = new List<PdfCore.PdfTableCellImage>();
        if (!options.IncludeImages) {
            return images;
        }

        foreach (RtfParagraph paragraph in cell.Paragraphs) {
            foreach (IRtfInline inline in paragraph.Inlines) {
                if (inline is RtfImage image && IsPdfSupportedImage(image) && image.Data.Length > 0) {
                    images.Add(new PdfCore.PdfTableCellImage(image.Data, GetImageWidth(image, options), GetImageHeight(image, options)));
                }
            }
        }

        return images;
    }

    private static bool IsPdfSupportedImage(RtfImage image) => image.Format == RtfImageFormat.Png || image.Format == RtfImageFormat.Jpeg;

    private static double GetImageWidth(RtfImage image, RtfPdfSaveOptions options) {
        if (image.DesiredWidthTwips.HasValue && image.DesiredWidthTwips.Value > 0) {
            return RtfPdfMapping.TwipsToPoints(image.DesiredWidthTwips.Value);
        }

        return options.DefaultImageWidth;
    }

    private static double GetImageHeight(RtfImage image, RtfPdfSaveOptions options) {
        if (image.DesiredHeightTwips.HasValue && image.DesiredHeightTwips.Value > 0) {
            return RtfPdfMapping.TwipsToPoints(image.DesiredHeightTwips.Value);
        }

        return options.DefaultImageHeight;
    }

    private static void ApplyMetadata(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        if (!options.IncludeMetadata) {
            return;
        }

        pdf.Meta(
            title: document.Info.Title,
            author: document.Info.Author,
            subject: document.Info.Subject,
            keywords: document.Info.Keywords);
    }

    private static void ApplyPageSetup(RtfPageSetup setup, PdfCore.PdfOptions options) {
        if (setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) {
            options.PageWidth = RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value);
        }

        if (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) {
            options.PageHeight = RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value);
        }

        if (setup.Landscape && options.PageWidth < options.PageHeight) {
            double width = options.PageWidth;
            options.PageWidth = options.PageHeight;
            options.PageHeight = width;
        }

        if (setup.MarginLeftTwips.HasValue) {
            options.MarginLeft = RtfPdfMapping.TwipsToPoints(setup.MarginLeftTwips.Value);
        }

        if (setup.MarginRightTwips.HasValue) {
            options.MarginRight = RtfPdfMapping.TwipsToPoints(setup.MarginRightTwips.Value);
        }

        if (setup.MarginTopTwips.HasValue) {
            options.MarginTop = RtfPdfMapping.TwipsToPoints(setup.MarginTopTwips.Value);
        }

        if (setup.MarginBottomTwips.HasValue) {
            options.MarginBottom = RtfPdfMapping.TwipsToPoints(setup.MarginBottomTwips.Value);
        }
    }
}
