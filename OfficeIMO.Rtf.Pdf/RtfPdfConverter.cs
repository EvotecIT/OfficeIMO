using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class RtfPdfConverter {
    internal static PdfCore.PdfDocument Convert(RtfDocument document, RtfPdfSaveOptions? options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        RtfPdfSaveOptions normalized = (options ?? new RtfPdfSaveOptions()).Normalize();
        PdfCore.PdfOptions pdfOptions = normalized.PdfOptions ?? new PdfCore.PdfOptions();
        ApplyPageSetup(document, document.PageSetup, pdfOptions);
        if (document.Sections.Count > 0) {
            ApplyPageSetup(document, document.Sections[0].PageSetup, pdfOptions);
        }

        ApplyHeaderFooters(document, pdfOptions, normalized);

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create(pdfOptions);
        ApplyMetadata(document, pdf, normalized);
        PdfRenderState state = new PdfRenderState(document);

        RenderDocumentBlocks(document, pdf, normalized, state, pdfOptions);

        RenderNotes(document, pdf, normalized, state);
        return pdf;
    }

    private static void RenderDocumentBlocks(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state, PdfCore.PdfOptions pdfOptions) {
        if (document.Sections.Count == 0) {
            RenderBlocks(document, document.Blocks, pdf, options, state);
            return;
        }

        for (int index = 0; index < document.Sections.Count; index++) {
            RtfSection section = document.Sections[index];
            if (index == 0) {
                RenderBlocks(document, section.Blocks, pdf, options, state);
                continue;
            }

            if (!StartsNewPdfPage(section.BreakKind)) {
                RenderBlocks(document, section.Blocks, pdf, options, state);
                continue;
            }

            pdf.Section(page => {
                ApplyPageSetup(document, section.PageSetup, page, pdfOptions);
                RenderBlocks(document, section.Blocks, pdf, options, state);

                while (index + 1 < document.Sections.Count && !StartsNewPdfPage(document.Sections[index + 1].BreakKind)) {
                    index++;
                    RenderBlocks(document, document.Sections[index].Blocks, pdf, options, state);
                }
            });
        }
    }

    private static void RenderBlocks(RtfDocument document, IEnumerable<IRtfBlock> blocks, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        foreach (IRtfBlock block in blocks) {
            RenderBlock(document, block, pdf, options, state);
        }
    }

    private static bool StartsNewPdfPage(RtfSectionBreakKind breakKind) {
        switch (breakKind) {
            case RtfSectionBreakKind.Continuous:
                return false;
            default:
                return true;
        }
    }

    private static void RenderBlock(RtfDocument document, IRtfBlock block, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        switch (block) {
            case RtfParagraph paragraph:
                RenderParagraph(document, paragraph, pdf, options, state);
                break;
            case RtfTable table when options.IncludeTables:
                RenderTable(document, table, pdf, options, state);
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

    private static void RenderParagraph(RtfDocument document, RtfParagraph paragraph, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        if (RtfPdfMapping.HasPageBreakBefore(document, paragraph)) {
            pdf.PageBreak();
        }

        PdfCore.PdfAlign align = RtfPdfMapping.ToPdfAlign(paragraph.Alignment);
        PdfCore.PdfParagraphStyle? style = RtfPdfMapping.ToPdfParagraphStyle(document, paragraph);
        List<PdfCore.TextRun> pendingRuns = new List<PdfCore.TextRun>();
        bool emitted = false;
        AppendListMarker(paragraph, pendingRuns, state);

        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, pendingRuns, options, state);
                    break;
                case RtfBreak rtfBreak when rtfBreak.Kind == RtfBreakKind.Page:
                    FlushParagraph(pdf, pendingRuns, align, style);
                    emitted = true;
                    pdf.PageBreak();
                    break;
                case RtfBreak:
                    pendingRuns.Add(PdfCore.TextRun.LineBreak());
                    break;
                case RtfField field:
                    AppendParagraphRuns(document, field.Result, pendingRuns, options, state);
                    break;
                case RtfImage image:
                    FlushParagraph(pdf, pendingRuns, align, style);
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
                    FlushParagraph(pdf, pendingRuns, align, style);
                    emitted = true;
                    pdf.Bookmark(marker.Name);
                    break;
            }
        }

        if (pendingRuns.Count > 0 || !emitted) {
            FlushParagraph(pdf, pendingRuns, align, style);
        }
    }

    private static void RenderTable(RtfDocument document, RtfTable table, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        List<PdfCore.PdfTableCell[]> rows = new List<PdfCore.PdfTableCell[]>();
        foreach (RtfTableRow row in table.Rows) {
            List<PdfCore.PdfTableCell> cells = new List<PdfCore.PdfTableCell>();
            foreach (RtfTableCell cell in row.Cells) {
                if (cell.HorizontalMerge == RtfTableCellMerge.Continue || cell.VerticalMerge == RtfTableCellMerge.Continue) {
                    continue;
                }

                List<PdfCore.TextRun> runs = BuildCellRuns(document, cell, options, state);
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

    private static void FlushParagraph(PdfCore.PdfDocument pdf, List<PdfCore.TextRun> runs, PdfCore.PdfAlign align, PdfCore.PdfParagraphStyle? style) {
        List<PdfCore.TextRun> snapshot = runs.Count == 0
            ? new List<PdfCore.TextRun> { PdfCore.TextRun.Normal(string.Empty) }
            : new List<PdfCore.TextRun>(runs);
        runs.Clear();
        pdf.Paragraph(paragraph => paragraph.Runs(snapshot), align, style: style);
    }

    private static void AppendParagraphRuns(RtfDocument document, RtfParagraph paragraph, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options, PdfRenderState state, bool collectNotes = true) {
        AppendListMarker(paragraph, runs, state);
        foreach (IRtfInline inline in paragraph.Inlines) {
            switch (inline) {
                case RtfRun run:
                    AppendRun(document, run, runs, options, state, collectNotes);
                    break;
                case RtfBreak:
                    runs.Add(PdfCore.TextRun.LineBreak());
                    break;
                case RtfField field:
                    AppendParagraphRuns(document, field.Result, runs, options, state, collectNotes);
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

    private static void AppendRun(RtfDocument document, RtfRun run, List<PdfCore.TextRun> runs, RtfPdfSaveOptions options, PdfRenderState state, bool collectNotes = true) {
        if (run.Hidden && !options.IncludeHiddenText) {
            return;
        }

        string text = run.Text ?? string.Empty;
        if (collectNotes && run.Note != null) {
            state.AddNote(run.Note, text);
        }

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

    private static void RenderNotes(RtfDocument document, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        if (!options.IncludeNotes || state.NoteReferences.Count == 0) {
            return;
        }

        pdf.HR(spacingBefore: 8D, spacingAfter: 4D);
        foreach (PdfNoteReference noteReference in state.NoteReferences) {
            List<PdfCore.TextRun> runs = new List<PdfCore.TextRun> {
                PdfCore.TextRun.Bolded(GetNoteLabel(noteReference), fontSize: 9D)
            };

            if (noteReference.Note.Paragraphs.Count == 0) {
                runs.Add(PdfCore.TextRun.Normal(string.Empty, fontSize: 9D));
            }

            for (int i = 0; i < noteReference.Note.Paragraphs.Count; i++) {
                if (i > 0) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                AppendParagraphRuns(document, noteReference.Note.Paragraphs[i], runs, options, state, collectNotes: false);
            }

            pdf.Paragraph(paragraph => paragraph.Runs(runs));
        }
    }

    private static string GetNoteLabel(PdfNoteReference reference) {
        string marker = string.IsNullOrWhiteSpace(reference.Marker)
            ? reference.Ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : reference.Marker.Trim();

        switch (reference.Note.Kind) {
            case RtfNoteKind.Endnote:
                return "Endnote " + marker + ": ";
            case RtfNoteKind.Annotation:
                string author = string.IsNullOrWhiteSpace(reference.Note.Author)
                    ? string.Empty
                    : " (" + reference.Note.Author!.Trim() + ")";
                return "Annotation " + marker + author + ": ";
            default:
                return "Footnote " + marker + ": ";
        }
    }

    private static void AppendPlainText(string text, List<PdfCore.TextRun> runs) {
        if (!string.IsNullOrEmpty(text)) {
            runs.Add(PdfCore.TextRun.Normal(text));
        }
    }

    private static List<PdfCore.TextRun> BuildCellRuns(RtfDocument document, RtfTableCell cell, RtfPdfSaveOptions options, PdfRenderState state) {
        List<PdfCore.TextRun> runs = new List<PdfCore.TextRun>();
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            if (i > 0) {
                runs.Add(PdfCore.TextRun.LineBreak());
            }

            AppendParagraphRuns(document, cell.Paragraphs[i], runs, options, state);
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

    private static void AppendListMarker(RtfParagraph paragraph, List<PdfCore.TextRun> runs, PdfRenderState state) {
        string? marker = GetListMarker(paragraph, state);
        if (marker != null && marker.Length > 0) {
            runs.Add(PdfCore.TextRun.Normal(marker));
        }
    }

    private static string? GetListMarker(RtfParagraph paragraph, PdfRenderState state) {
        if (paragraph.ListKind == RtfListKind.None) {
            return null;
        }

        if (paragraph.ListText != null) {
            string markerText = NormalizeListMarkerText(paragraph.ListText.ToPlainText());
            if (paragraph.ListKind == RtfListKind.Decimal) {
                state.AdvanceDecimalList(paragraph, markerText);
            }

            return EnsureMarkerSeparator(markerText);
        }

        if (paragraph.ListKind == RtfListKind.Bullet) {
            return "\u2022 ";
        }

        return state.NextDecimalMarker(paragraph).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". ";
    }

    private static string NormalizeListMarkerText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return text
            .Replace("\r\n", " ")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\f', ' ')
            .Replace('\v', ' ')
            .Replace('\t', ' ')
            .Trim();
    }

    private static string EnsureMarkerSeparator(string marker) {
        if (marker.Length == 0 || char.IsWhiteSpace(marker[marker.Length - 1])) {
            return marker;
        }

        return marker + " ";
    }

    private static bool TryReadLeadingIntegerMarker(string marker, out int value) {
        value = 0;
        int index = 0;
        while (index < marker.Length && char.IsWhiteSpace(marker[index])) {
            index++;
        }

        int start = index;
        while (index < marker.Length && char.IsDigit(marker[index])) {
            int digit = marker[index] - '0';
            if (value > (int.MaxValue - digit) / 10) {
                value = 0;
                return false;
            }

            value = (value * 10) + digit;
            index++;
        }

        return index > start;
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

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfOptions options) {
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

        if (setup.PageNumberStart.HasValue) {
            options.PageNumberStart = setup.PageNumberStart.Value;
        }

        if (setup.PageNumberFormat.HasValue) {
            options.PageNumberStyle = RtfPdfMapping.ToPdfPageNumberStyle(setup.PageNumberFormat.Value);
        }

        PdfCore.PdfPageBorder? border = RtfPdfMapping.ToPdfPageBorder(document, setup.PageBorders);
        if (border != null) {
            options.PageBorder = border;
        }
    }

    private static void ApplyPageSetup(RtfDocument document, RtfPageSetup setup, PdfCore.PdfPageCompose page, PdfCore.PdfOptions inheritedOptions) {
        double width = setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0
            ? RtfPdfMapping.TwipsToPoints(setup.PaperWidthTwips.Value)
            : inheritedOptions.PageWidth;
        double height = setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0
            ? RtfPdfMapping.TwipsToPoints(setup.PaperHeightTwips.Value)
            : inheritedOptions.PageHeight;

        if (setup.Landscape && width < height) {
            double swap = width;
            width = height;
            height = swap;
        }

        if ((setup.PaperWidthTwips.HasValue && setup.PaperWidthTwips.Value > 0) ||
            (setup.PaperHeightTwips.HasValue && setup.PaperHeightTwips.Value > 0) ||
            setup.Landscape) {
            page.Size(width, height);
        }

        if (HasAnyMargin(setup)) {
            page.Margin(
                setup.MarginLeftTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginLeftTwips.Value) : inheritedOptions.MarginLeft,
                setup.MarginTopTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginTopTwips.Value) : inheritedOptions.MarginTop,
                setup.MarginRightTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginRightTwips.Value) : inheritedOptions.MarginRight,
                setup.MarginBottomTwips.HasValue ? RtfPdfMapping.TwipsToPoints(setup.MarginBottomTwips.Value) : inheritedOptions.MarginBottom);
        }

        if (setup.PageNumberStart.HasValue) {
            page.PageNumberStart(setup.PageNumberStart.Value);
        }

        if (setup.PageNumberFormat.HasValue) {
            page.PageNumberStyle(RtfPdfMapping.ToPdfPageNumberStyle(setup.PageNumberFormat.Value));
        }

        PdfCore.PdfPageBorder? border = RtfPdfMapping.ToPdfPageBorder(document, setup.PageBorders);
        if (border != null) {
            page.PageBorder(border);
        }
    }

    private static bool HasAnyMargin(RtfPageSetup setup) {
        return setup.MarginLeftTwips.HasValue ||
               setup.MarginRightTwips.HasValue ||
               setup.MarginTopTwips.HasValue ||
               setup.MarginBottomTwips.HasValue;
    }

    private static void ApplyHeaderFooters(RtfDocument document, PdfCore.PdfOptions options, RtfPdfSaveOptions saveOptions) {
        if (!saveOptions.IncludeHeaderFooters || document.HeaderFooters.Count == 0) {
            return;
        }

        string? defaultHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.RightHeader)
            ?? GetHeaderFooterText(document, RtfHeaderFooterKind.Header);
        if (defaultHeader != null && defaultHeader.Length > 0) {
            options.ShowHeader = true;
            options.HeaderFormat = defaultHeader;
        }

        string? defaultFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.RightFooter)
            ?? GetHeaderFooterText(document, RtfHeaderFooterKind.Footer);
        if (defaultFooter != null && defaultFooter.Length > 0) {
            options.ShowPageNumbers = true;
            options.FooterFormat = defaultFooter;
        }

        string? firstHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.FirstHeader);
        string? firstFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.FirstFooter);
        if ((firstHeader != null && firstHeader.Length > 0) ||
            (firstFooter != null && firstFooter.Length > 0) ||
            document.PageSetup.DifferentFirstPageHeaderFooter) {
            options.DifferentFirstPageHeaderFooter = true;
            if (firstHeader != null && firstHeader.Length > 0) {
                options.FirstPageHeaderFormat = firstHeader;
            }

            if (firstFooter != null && firstFooter.Length > 0) {
                options.FirstPageFooterFormat = firstFooter;
            }
        }

        string? evenHeader = GetHeaderFooterText(document, RtfHeaderFooterKind.LeftHeader);
        string? evenFooter = GetHeaderFooterText(document, RtfHeaderFooterKind.LeftFooter);
        if ((evenHeader != null && evenHeader.Length > 0) ||
            (evenFooter != null && evenFooter.Length > 0)) {
            options.DifferentOddAndEvenPagesHeaderFooter = true;
            if (evenHeader != null && evenHeader.Length > 0) {
                options.EvenPageHeaderFormat = evenHeader;
            }

            if (evenFooter != null && evenFooter.Length > 0) {
                options.EvenPageFooterFormat = evenFooter;
            }
        }
    }

    private static string? GetHeaderFooterText(RtfDocument document, RtfHeaderFooterKind kind) {
        RtfHeaderFooter? headerFooter = document.HeaderFooters.FirstOrDefault(item => item.Kind == kind);
        if (headerFooter == null) {
            return null;
        }

        string text = NormalizeHeaderFooterText(headerFooter.ToPlainText());
        return text.Length == 0 ? null : text;
    }

    private static string NormalizeHeaderFooterText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return text
            .Replace("\r\n", " ")
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\f', ' ')
            .Replace('\v', ' ')
            .Trim();
    }

    private sealed class PdfRenderState {
        private readonly RtfDocument _document;
        private readonly Dictionary<string, int> _listCounters = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<PdfNoteReference> _noteReferences = new List<PdfNoteReference>();

        public PdfRenderState(RtfDocument document) {
            _document = document;
        }

        public IReadOnlyList<PdfNoteReference> NoteReferences => _noteReferences.AsReadOnly();

        public void AddNote(RtfNote note, string marker) {
            _noteReferences.Add(new PdfNoteReference(note, marker, _noteReferences.Count + 1));
        }

        public int NextDecimalMarker(RtfParagraph paragraph) {
            string key = GetListCounterKey(paragraph);
            if (!_listCounters.TryGetValue(key, out int value)) {
                value = GetListStart(paragraph);
            }

            _listCounters[key] = value + 1;
            return value;
        }

        public void AdvanceDecimalList(RtfParagraph paragraph, string markerText) {
            if (TryReadLeadingIntegerMarker(markerText, out int value)) {
                _listCounters[GetListCounterKey(paragraph)] = value + 1;
                return;
            }

            NextDecimalMarker(paragraph);
        }

        private int GetListStart(RtfParagraph paragraph) {
            int levelIndex = paragraph.ListLevel ?? 0;
            RtfListOverride? listOverride = paragraph.ListId.HasValue
                ? _document.ListOverrides.FirstOrDefault(item => item.Id == paragraph.ListId.Value)
                : null;
            RtfListLevelOverride? levelOverride = listOverride?.LevelOverrides.ElementAtOrDefault(levelIndex);
            if (levelOverride?.StartAt.HasValue == true) {
                return levelOverride.StartAt.Value;
            }

            int? definitionId = paragraph.ListDefinitionId ?? listOverride?.ListId;
            RtfListDefinition? definition = definitionId.HasValue
                ? _document.ListDefinitions.FirstOrDefault(item => item.Id == definitionId.Value)
                : null;
            RtfListLevel? level = definition?.Levels.FirstOrDefault(item => item.LevelIndex == levelIndex)
                ?? definition?.Levels.ElementAtOrDefault(levelIndex);
            return level?.StartAt ?? 1;
        }

        private static string GetListCounterKey(RtfParagraph paragraph) {
            int listId = paragraph.ListId ?? paragraph.ListDefinitionId ?? 0;
            int level = paragraph.ListLevel ?? 0;
            return listId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" +
                   level.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private sealed class PdfNoteReference {
        public PdfNoteReference(RtfNote note, string marker, int ordinal) {
            Note = note;
            Marker = marker;
            Ordinal = ordinal;
        }

        public RtfNote Note { get; }

        public string Marker { get; }

        public int Ordinal { get; }
    }
}
