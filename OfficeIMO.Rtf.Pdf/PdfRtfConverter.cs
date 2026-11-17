using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static class PdfRtfConverter {
    private const int HeadingStyleBaseId = 1;

    public static RtfDocument Convert(PdfCore.PdfDocumentReadResult source, PdfRtfExportOptions? options) {
        ArgumentNullException.ThrowIfNull(source);

        PdfRtfExportOptions readOptions = options ?? new PdfRtfExportOptions();
        readOptions.CancellationToken.ThrowIfCancellationRequested();
        PdfCore.PdfLogicalContentProjection.RequireComplete(source, readOptions.CancellationToken);
        RtfDocument document = RtfDocument.Create();
        var runStyles = new RunStyleCatalog(document);
        var lists = new SourceListNumbering(document);
        AddUnsupportedContentDiagnostics(source, readOptions);

        if (readOptions.IncludeMetadata) {
            CopyMetadata(source.Metadata, document.Properties);
        }

        if (source.Pages.Count > 0) {
            ApplyFirstPageSetup(source.Pages[0], document);
        }

        if (readOptions.CreateHeadingStyles && readOptions.ImportHeadings) {
            AddHeadingStyles(document);
        }

        for (int pageIndex = 0; pageIndex < source.Pages.Count; pageIndex++) {
            readOptions.CancellationToken.ThrowIfCancellationRequested();
            ImportPage(source.Pages[pageIndex], document, readOptions, runStyles, lists, pageIndex > 0 && readOptions.PreservePageBreaks);
        }

        return document;
    }

    private static void AddUnsupportedContentDiagnostics(
        PdfCore.PdfDocumentReadResult source,
        PdfRtfExportOptions options) {
        AddOmissionWarning(options, "TABLES_NOT_IMPORTED", "Table", source.Tables.Count,
            "Detected PDF tables are not reconstructed by the semantic RTF importer.");
        AddOmissionWarning(options, "IMAGES_NOT_IMPORTED", "Image", source.Images.Count,
            "Detected PDF images are not embedded by the semantic RTF importer.");
        AddOmissionWarning(options, "LINKS_NOT_IMPORTED", "LinkAnnotation", source.Links.Count,
            "PDF link annotations are not reconstructed by the semantic RTF importer.");
        AddOmissionWarning(options, "FORM_WIDGETS_NOT_IMPORTED", "FormWidget", source.FormWidgets.Count,
            "PDF form widgets are not reconstructed by the semantic RTF importer.");
    }

    private static void AddOmissionWarning(
        PdfRtfExportOptions options,
        string code,
        string source,
        int count,
        string message) {
        if (count <= 0) {
            return;
        }

        options.Report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Rtf.Pdf",
            code,
            source,
            message,
            PdfCore.PdfConversionWarningSeverity.Warning,
            global::OfficeIMO.OfficeConversionLossKind.Omission,
            details: new Dictionary<string, string> {
                ["count"] = count.ToString(System.Globalization.CultureInfo.InvariantCulture)
            }));
    }

    private static void ImportPage(
        PdfCore.PdfLogicalPage page,
        RtfDocument document,
        PdfRtfExportOptions options,
        RunStyleCatalog runStyles,
        SourceListNumbering lists,
        bool pageBreakBeforeFirstParagraph) {
        var consumed = new HashSet<PdfCore.PdfLogicalTextBlock>();
        bool emittedPageContent = false;
        bool pendingPageBreak = pageBreakBeforeFirstParagraph;

        foreach (PdfCore.PdfLogicalTextBlock block in page.TextBlocks) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (consumed.Contains(block)) {
                continue;
            }

            RtfParagraph? paragraph = null;
            if (options.ImportHeadings && TryFindHeading(page, block, out PdfCore.PdfLogicalHeading? heading)) {
                paragraph = AddHeading(document, heading!, runStyles);
                consumed.Add(block);
            } else if (options.ImportLists && TryFindListItem(page, block, out PdfCore.PdfLogicalListItem? listItem)) {
                paragraph = AddListItem(document, listItem!, runStyles, lists);
                consumed.Add(block);
            } else if (TryFindParagraph(page, block, out PdfCore.PdfLogicalParagraph? logicalParagraph)) {
                paragraph = AddParagraph(document, logicalParagraph!, runStyles);
                foreach (PdfCore.PdfLogicalTextBlock line in logicalParagraph!.Lines) {
                    consumed.Add(line);
                }
            } else if (!string.IsNullOrWhiteSpace(block.Text)) {
                paragraph = AddTextBlock(document, block, runStyles);
                consumed.Add(block);
            }

            if (paragraph is null) {
                continue;
            }

            // Anything written between two list items ends the list; the next item starts a new one.
            if (paragraph.ListKind == RtfListKind.None) lists.End();
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

    private static RtfParagraph AddHeading(
        RtfDocument document,
        PdfCore.PdfLogicalHeading heading,
        RunStyleCatalog runStyles) {
        int level = Math.Max(1, Math.Min(3, heading.Level));
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetStyle(HeadingStyleBaseId + level - 1);
        paragraph.SetOutlineLevel(level - 1);
        paragraph.SetParagraphSpacing(afterTwips: 120);
        paragraph.SetPagination(keepWithNext: true);

        AppendStyledLine(paragraph, heading.Line, runStyles);
        foreach (RtfRun run in paragraph.Runs) run.SetBold();

        return paragraph;
    }

    private static RtfParagraph AddListItem(
        RtfDocument document,
        PdfCore.PdfLogicalListItem item,
        RunStyleCatalog runStyles,
        SourceListNumbering lists) {
        PdfCore.PdfLogicalListMarker marker = item.MarkerInfo;
        RtfListKind kind = marker.IsOrdered ? RtfListKind.Decimal : RtfListKind.Bullet;
        int level = Math.Max(0, Math.Min(8, item.Level - 1));
        int listId = lists.Assign(level, marker);
        RtfParagraph paragraph = document.AddParagraph();
        AppendStyledRuns(paragraph, item.Runs, item.Text, runStyles);
        if (item.IsContentsEntry && !string.IsNullOrEmpty(item.TargetPageLabel)) paragraph.AddText(" \u2014 " + item.TargetPageLabel);
        paragraph.SetParagraphSpacing(afterTwips: 120);
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

    private static RtfParagraph AddParagraph(
        RtfDocument document,
        PdfCore.PdfLogicalParagraph source,
        RunStyleCatalog runStyles) {
        RtfParagraph paragraph = document.AddParagraph();
        for (int lineIndex = 0; lineIndex < source.Lines.Count; lineIndex++) {
            if (lineIndex > 0) paragraph.AddText(" ");
            AppendStyledLine(paragraph, source.Lines[lineIndex], runStyles);
        }
        paragraph.SetParagraphSpacing(afterTwips: 120);
        return paragraph;
    }

    private static RtfParagraph AddTextBlock(
        RtfDocument document,
        PdfCore.PdfLogicalTextBlock block,
        RunStyleCatalog runStyles) {
        RtfParagraph paragraph = document.AddParagraph();
        AppendStyledLine(paragraph, block, runStyles);
        paragraph.SetParagraphSpacing(afterTwips: 120);
        return paragraph;
    }

    private static void AppendStyledLine(
        RtfParagraph paragraph,
        PdfCore.PdfLogicalTextBlock line,
        RunStyleCatalog runStyles) {
        AppendStyledRuns(paragraph, line.Runs, line.Text, runStyles);
    }

    private static void AppendStyledRuns(
        RtfParagraph paragraph,
        IReadOnlyList<PdfCore.PdfLogicalTextRun> runs,
        string fallbackText,
        RunStyleCatalog runStyles) {
        if (runs.Count == 0) {
            paragraph.AddText(fallbackText);
            return;
        }

        for (int i = 0; i < runs.Count; i++) {
            PdfCore.PdfLogicalTextRun source = runs[i];
            RtfRun run = paragraph.AddText(source.Text);
            if (source.IsBold) run.SetBold();
            if (source.IsItalic) run.SetItalic();
            if (source.FontSize > 0D) run.SetFontSize(source.FontSize);
            if (source.Color.HasValue && source.Color.Value.A > 0) {
                run.SetForegroundColor(runStyles.GetColorIndex(source.Color.Value));
            }
        }
    }

    private static void CopyMetadata(PdfCore.PdfDocumentProperties source, RtfDocumentInfo target) {
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

    /// <summary>
    /// Gives each source list its own RTF list, so a second list starts again instead of counting on from the first.
    /// A level that changes its numbering style or punctuation, or a number that does not follow on, starts a new
    /// list. Each level of the list definition takes its format (bullet, decimal, letter or Roman, with the source
    /// punctuation) and start value from the first item drawn at that level.
    /// </summary>
    private sealed class SourceListNumbering {
        private const int MaximumLevels = 9;
        private readonly RtfDocument _document;
        private readonly PdfCore.PdfLogicalListMarker?[] _markers = new PdfCore.PdfLogicalListMarker?[MaximumLevels];
        private readonly int[] _starts = new int[MaximumLevels];
        private readonly int?[] _counters = new int?[MaximumLevels];
        private readonly bool[] _carried = new bool[MaximumLevels];
        private RtfListDefinition? _definition;
        private int _nextId = 1;
        private int _currentId;

        internal SourceListNumbering(RtfDocument document) {
            _document = document;
        }

        internal void End() => _currentId = 0;

        internal int Assign(int level, PdfCore.PdfLogicalListMarker marker) {
            if (_currentId == 0) Start(carriedLevels: 0);
            else if (!Continues(level, marker)) Start(carriedLevels: level);
            if (_markers[level] == null) {
                SeedParents(level, marker);
                ConfigureLevel(level, marker, Math.Max(0, marker.Ordinal ?? 1));
            }
            _counters[level] = marker.Ordinal ?? (_counters[level] + 1 ?? _starts[level]);
            for (int deeper = level + 1; deeper < MaximumLevels; deeper++) _counters[deeper] = null;
            return _currentId;
        }

        private bool Continues(int level, PdfCore.PdfLogicalListMarker marker) {
            // A parent level carried into a restarted list only shows its number; its next item needs a list that
            // counts on from there.
            if (_carried[level]) return false;
            // "4.3" after "3.2" belongs under another parent, which the list has to show.
            if (!MatchesParents(level, marker)) return false;
            if (_markers[level] is not PdfCore.PdfLogicalListMarker current) return true;
            if (current.Format != marker.Format || current.Prefix != marker.Prefix || current.Suffix != marker.Suffix ||
                current.IncludesParentNumbers != marker.IncludesParentNumbers) return false;
            return !marker.Ordinal.HasValue || marker.Ordinal.Value == (_counters[level] + 1 ?? _starts[level]);
        }

        /// <summary>
        /// Starts a new list. A nested list that restarts keeps its parent levels, at their current numbers, so
        /// "2.5" after "2." still shows the 2.
        /// </summary>
        private void Start(int carriedLevels) {
            var parents = new (PdfCore.PdfLogicalListMarker Marker, int Number)?[carriedLevels];
            for (int level = 0; level < carriedLevels; level++) {
                if (_markers[level] is PdfCore.PdfLogicalListMarker marker) parents[level] = (marker, _counters[level] ?? _starts[level]);
            }

            _currentId = _nextId++;
            _definition = _document.AddListDefinition(_currentId, "PDF list");
            _document.AddListOverride(_currentId, _currentId);
            for (int level = 0; level < MaximumLevels; level++) {
                RtfListLevel listLevel = _definition.AddLevel(RtfListKind.Bullet);
                listLevel.NumberFormat = 23;
                listLevel.Text = "\u2022";
                listLevel.LeftIndentTwips = 720 + level * 360;
                listLevel.FirstLineIndentTwips = -360;
                _markers[level] = null;
                _starts[level] = 1;
                _counters[level] = null;
                _carried[level] = false;
            }

            for (int level = 0; level < carriedLevels; level++) {
                if (parents[level] is not { } parent) continue;
                ConfigureLevel(level, parent.Marker, parent.Number);
                _counters[level] = parent.Number;
                _carried[level] = true;
            }
        }

        private bool MatchesParents(int level, PdfCore.PdfLogicalListMarker marker) {
            if (!marker.IncludesParentNumbers || marker.ParentNumbers.Count != level) return true;
            for (int parent = 0; parent < level; parent++) {
                if (_counters[parent] is int number && number != marker.ParentNumbers[parent]) return false;
            }
            return true;
        }

        /// <summary>An outline number with no parent item before it ("2.5" first) sets up the parents it shows.</summary>
        private void SeedParents(int level, PdfCore.PdfLogicalListMarker marker) {
            if (!marker.IncludesParentNumbers || marker.ParentNumbers.Count != level) return;
            for (int parent = 0; parent < level; parent++) {
                int number = marker.ParentNumbers[parent];
                // A parent carried into a restarted list takes the number the outline item shows.
                if (_markers[parent] != null && (!_carried[parent] || _counters[parent] == number)) continue;
                ConfigureLevel(parent, _markers[parent] ?? PdfCore.PdfLogicalListMarker.Parse(number.ToString(System.Globalization.CultureInfo.InvariantCulture) + "."), number);
                _counters[parent] = number;
                _carried[parent] = true;
            }
        }

        private void ConfigureLevel(int level, PdfCore.PdfLogicalListMarker marker, int start) {
            _markers[level] = marker;
            if (!marker.IsOrdered) return;
            RtfListLevel listLevel = _definition!.Levels[level];
            listLevel.Kind = RtfListKind.Decimal;
            listLevel.NumberFormat = marker.Format switch {
                PdfCore.PdfLogicalListNumberFormat.UpperRoman => 1,
                PdfCore.PdfLogicalListNumberFormat.LowerRoman => 2,
                PdfCore.PdfLogicalListNumberFormat.UpperLetter => 3,
                PdfCore.PdfLogicalListNumberFormat.LowerLetter => 4,
                _ => 0
            };
            string suffix = marker.Suffix.Length == 0 && marker.Prefix.Length == 0 ? "." : marker.Suffix;
            var levelText = new System.Text.StringBuilder(marker.Prefix);
            // "1.2" repeats its parent numbers, written as "%1.%2".
            for (int shown = marker.IncludesParentNumbers ? 0 : level; shown <= level; shown++) {
                levelText.Append('%').Append((shown + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (shown < level) levelText.Append('.');
            }
            listLevel.Text = levelText.Append(suffix).ToString();
            _starts[level] = start;
            listLevel.StartAt = start;
            // "1.1" under a Roman or lettered parent still shows the parent as a number.
            if (marker.IncludesParentNumbers && marker.Format == PdfCore.PdfLogicalListNumberFormat.Number) listLevel.LegalNumbering = true;
        }
    }

    private static int PointsToTwips(double points) {
        if (points <= 0 || double.IsNaN(points) || double.IsInfinity(points)) {
            return 0;
        }

        return (int)Math.Round(points * 20D, MidpointRounding.AwayFromZero);
    }

    private sealed class RunStyleCatalog {
        private readonly RtfDocument _document;
        private readonly Dictionary<OfficeIMO.OfficeColor, int> _colors =
            new Dictionary<OfficeIMO.OfficeColor, int>();

        internal RunStyleCatalog(RtfDocument document) {
            _document = document;
        }

        internal int GetColorIndex(OfficeIMO.OfficeColor color) {
            if (_colors.TryGetValue(color, out int index)) {
                return index;
            }

            index = _document.AddColor(color.R, color.G, color.B);
            _colors[color] = index;
            return index;
        }
    }
}
