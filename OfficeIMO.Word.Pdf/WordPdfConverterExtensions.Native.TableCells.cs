using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static List<PdfCore.PdfCellVerticalAlign>? CreateNativeTableVerticalAlignments(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var alignments = new List<PdfCore.PdfCellVerticalAlign>(columnCount);
            bool hasExplicitAlignment = false;
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                PdfCore.PdfCellVerticalAlign? columnAlignment = null;
                bool conflict = false;
                foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout)) {
                    if (columnIndex < cell.Column || columnIndex >= cell.Column + cell.ColumnSpan) {
                        continue;
                    }

                    PdfCore.PdfCellVerticalAlign alignment = MapNativeCellVerticalAlign(cell.Cell.VerticalAlignment);
                    if (columnAlignment == null) {
                        columnAlignment = alignment;
                    } else if (columnAlignment.Value != alignment) {
                        conflict = true;
                        break;
                    }
                }

                PdfCore.PdfCellVerticalAlign resolved = conflict ? PdfCore.PdfCellVerticalAlign.Top : columnAlignment ?? PdfCore.PdfCellVerticalAlign.Top;
                if (resolved != PdfCore.PdfCellVerticalAlign.Top) {
                    hasExplicitAlignment = true;
                }

                alignments.Add(resolved);
            }

            return hasExplicitAlignment ? alignments : null;
        }

        private static int GetNativeTableColumnCount(TableLayout layout) {
            if (layout.ColumnWidths.Length > 0) {
                return layout.ColumnWidths.Length;
            }

            int columnCount = 0;
            foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
                int logicalColumn = 0;
                foreach (WordTableCell cell in row) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    logicalColumn += GetNativeCellColumnSpan(cell);
                }

                if (logicalColumn > columnCount) {
                    columnCount = logicalColumn;
                }
            }

            return columnCount;
        }

        private static IEnumerable<(WordTableCell Cell, int Column, int ColumnSpan)> EnumerateNativeTableCells(TableLayout layout) {
            foreach (IReadOnlyList<WordTableCell> row in layout.Rows) {
                int logicalColumn = 0;
                foreach (WordTableCell cell in row) {
                    if (IsNativeHorizontalMergeContinuation(cell)) {
                        continue;
                    }

                    int columnSpan = GetNativeCellColumnSpan(cell);
                    if (IsNativeVerticalMergeContinuation(cell)) {
                        logicalColumn += columnSpan;
                        continue;
                    }

                    yield return (cell, logicalColumn, columnSpan);
                    logicalColumn += columnSpan;
                }
            }
        }

        private static bool IsNativeHorizontalMergeContinuation(WordTableCell cell) =>
            cell.HorizontalMerge == W.MergedCellValues.Continue;

        private static bool IsNativeVerticalMergeContinuation(WordTableCell cell) =>
            cell.VerticalMerge == W.MergedCellValues.Continue;

        private static int GetNativeCellColumnSpan(WordTableCell cell) =>
            Math.Max(1, cell.ColumnSpan);

        private static int GetNativeCellRowSpan(WordTableCell cell) =>
            Math.Max(1, cell.RowSpan);

        private static (string? LinkUri, string? LinkContents) GetNativeCellLink(WordTableCell cell) {
            string? linkUri = null;
            string? linkContents = null;
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                if (!TryAddNativeCellLink(paragraph, ref linkUri, ref linkContents)) {
                    return (null, null);
                }

                foreach (WordParagraph run in paragraph.GetRuns()) {
                    if (!TryAddNativeCellLink(run, ref linkUri, ref linkContents)) {
                        return (null, null);
                    }
                }
            }

            return (linkUri, linkContents);
        }

        private static bool TryAddNativeCellLink(WordParagraph paragraph, ref string? linkUri, ref string? linkContents) {
            if (!paragraph.IsHyperLink || paragraph.Hyperlink == null) {
                return true;
            }

            Uri? uri = paragraph.Hyperlink.Uri;
            if (uri == null || !uri.IsAbsoluteUri) {
                return true;
            }

            string candidateUri = uri.AbsoluteUri;
            if (!string.IsNullOrEmpty(linkUri) && !string.Equals(linkUri, candidateUri, StringComparison.Ordinal)) {
                return false;
            }

            linkUri = candidateUri;
            string? contents = string.IsNullOrWhiteSpace(paragraph.Hyperlink.Tooltip)
                ? GetNativeCellParagraphText(paragraph)
                : paragraph.Hyperlink.Tooltip;
            linkContents ??= string.IsNullOrWhiteSpace(contents) ? null : contents;
            return true;
        }

        private static PdfCore.PdfCellBorder? CreateNativeTableCellBorder(WordTableCellBorder borders) {
            bool top = HasNativeBorder(borders.TopStyle);
            bool bottom = HasNativeBorder(borders.BottomStyle);
            bool left = HasNativeBorder(borders.LeftStyle);
            bool right = HasNativeBorder(borders.RightStyle);
            bool diagonalDown = HasNativeBorder(borders.TopLeftToBottomRightStyle);
            bool diagonalUp = HasNativeBorder(borders.TopRightToBottomLeftStyle);
            if (!top && !bottom && !left && !right && !diagonalDown && !diagonalUp) {
                return null;
            }

            if (!diagonalDown && !diagonalUp && TryGetNativeUniformTableCellBorder(borders, out PdfCore.PdfColor uniformColor, out double uniformWidth, out OfficeIMO.Drawing.OfficeStrokeDashStyle uniformDashStyle, out PdfCore.PdfCellBorderLineStyle uniformLineStyle)) {
                return new PdfCore.PdfCellBorder {
                    Color = uniformColor,
                    Width = uniformWidth,
                    DashStyle = uniformDashStyle,
                    LineStyle = uniformLineStyle,
                    Top = top,
                    Bottom = bottom,
                    Left = left,
                    Right = right
                };
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                Width = 0,
                TopBorder = CreateNativeCellBorderSide(borders.TopStyle, borders.TopColorHex, borders.TopSize),
                BottomBorder = CreateNativeCellBorderSide(borders.BottomStyle, borders.BottomColorHex, borders.BottomSize),
                LeftBorder = CreateNativeCellBorderSide(borders.LeftStyle, borders.LeftColorHex, borders.LeftSize),
                RightBorder = CreateNativeCellBorderSide(borders.RightStyle, borders.RightColorHex, borders.RightSize),
                DiagonalDownBorder = CreateNativeCellBorderSide(borders.TopLeftToBottomRightStyle, borders.TopLeftToBottomRightColorHex, borders.TopLeftToBottomRightSize),
                DiagonalUpBorder = CreateNativeCellBorderSide(borders.TopRightToBottomLeftStyle, borders.TopRightToBottomLeftColorHex, borders.TopRightToBottomLeftSize),
                Top = top,
                Bottom = bottom,
                Left = left,
                Right = right,
                DiagonalDown = diagonalDown,
                DiagonalUp = diagonalUp
            };
        }

        private static bool TryGetNativeUniformTableCellBorder(WordTableCellBorder borders, out PdfCore.PdfColor color, out double width, out OfficeIMO.Drawing.OfficeStrokeDashStyle dashStyle, out PdfCore.PdfCellBorderLineStyle lineStyle) {
            color = PdfCore.PdfColor.Black;
            width = 1D;
            dashStyle = OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            lineStyle = PdfCore.PdfCellBorderLineStyle.Standard;

            string? firstColor = null;
            uint? firstSize = null;
            W.BorderValues? firstStyle = null;
            bool hasFirst = false;
            foreach ((W.BorderValues? BorderStyle, string? Color, DocumentFormat.OpenXml.UInt32Value? Size) side in GetNativeTableCellBorderSides(borders)) {
                if (!HasNativeBorder(side.BorderStyle)) {
                    continue;
                }

                string? sideColor = NormalizeNativeBorderColor(side.Color);
                uint sideSize = side.Size?.Value ?? 4U;
                if (!hasFirst) {
                    firstColor = sideColor;
                    firstSize = sideSize;
                    firstStyle = side.BorderStyle;
                    hasFirst = true;
                    continue;
                }

                if (!string.Equals(firstColor, sideColor, StringComparison.OrdinalIgnoreCase) ||
                    firstSize.GetValueOrDefault() != sideSize ||
                    firstStyle != side.BorderStyle) {
                    return false;
                }
            }

            color = ParseNativeColor(firstColor) ?? PdfCore.PdfColor.Black;
            width = (firstSize ?? 4U) / 8D;
            dashStyle = ToNativeBorderDashStyle(firstStyle);
            lineStyle = ToNativeBorderLineStyle(firstStyle);
            return true;
        }

        private static IEnumerable<(W.BorderValues? BorderStyle, string? Color, DocumentFormat.OpenXml.UInt32Value? Size)> GetNativeTableCellBorderSides(WordTableCellBorder borders) {
            yield return (borders.TopStyle, borders.TopColorHex, borders.TopSize);
            yield return (borders.BottomStyle, borders.BottomColorHex, borders.BottomSize);
            yield return (borders.LeftStyle, borders.LeftColorHex, borders.LeftSize);
            yield return (borders.RightStyle, borders.RightColorHex, borders.RightSize);
        }

        private static PdfCore.PdfCellBorderSide? CreateNativeCellBorderSide(W.BorderValues? borderStyle, string? color, DocumentFormat.OpenXml.UInt32Value? size) {
            if (!HasNativeBorder(borderStyle)) {
                return null;
            }

            return new PdfCore.PdfCellBorderSide {
                Color = ParseNativeColor(NormalizeNativeBorderColor(color)) ?? PdfCore.PdfColor.Black,
                Width = (size?.Value ?? 4U) / 8D,
                DashStyle = ToNativeBorderDashStyle(borderStyle),
                LineStyle = ToNativeBorderLineStyle(borderStyle)
            };
        }

        private static PdfCore.PdfCellBorderSide? CreateNativeCellBorderSide(W.BorderType? border) =>
            border == null
                ? null
                : CreateNativeCellBorderSide(border.Val?.Value, border.Color?.Value, border.Size);

        private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToNativeBorderDashStyle(W.BorderValues? borderStyle) {
            string value = borderStyle?.ToString() ?? string.Empty;
            if (value.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0 &&
                value.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.DashDot;
            }

            if (value.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
            }

            if (value.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
            }

            return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
        }

        private static PdfCore.PdfCellBorderLineStyle ToNativeBorderLineStyle(W.BorderValues? borderStyle) =>
            borderStyle == W.BorderValues.Double
                ? PdfCore.PdfCellBorderLineStyle.TwoLine
                : PdfCore.PdfCellBorderLineStyle.Standard;

        private static string GetNativeCellText(WordTableCell cell) =>
            GetNativeCellText(cell, null);

        private readonly record struct NativeCellText(IReadOnlyList<PdfCore.TextRun> Runs, IReadOnlyList<PdfCore.PdfTableCellParagraph> Paragraphs);

        private static IReadOnlyList<PdfCore.TextRun> CreateNativeCellRuns(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById) {
            return CreateNativeCellText(cell, footnoteNumbersById, NativeDocumentDefaults.WordDefault, NativeTableStyleDefaults.Empty).Runs;
        }

        private static NativeCellText CreateNativeCellText(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById, NativeDocumentDefaults nativeDefaults) {
            return CreateNativeCellText(cell, footnoteNumbersById, nativeDefaults, NativeTableStyleDefaults.Empty);
        }

        private static NativeCellText CreateNativeCellText(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById, NativeDocumentDefaults nativeDefaults, NativeTableStyleDefaults tableStyleDefaults) {
            var runs = new List<PdfCore.TextRun>();
            var paragraphs = new List<PdfCore.PdfTableCellParagraph>();
            double? pendingSpacingAfter = null;
            List<WordParagraph> cellParagraphs = GetNativeCellParagraphs(cell).ToList();
            for (int i = 0; i < cellParagraphs.Count; i++) {
                WordParagraph paragraph = cellParagraphs[i];
                List<PdfCore.TextRun> paragraphRuns = CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById, tableStyleDefaults, nativeDefaults);
                if (paragraphRuns.Count == 0) {
                    continue;
                }

                if (runs.Count > 0) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                runs.AddRange(paragraphRuns);
                double spacingBefore = GetNativeCellParagraphSpacingBefore(paragraph, nativeDefaults, tableStyleDefaults);
                if (pendingSpacingAfter.HasValue) {
                    spacingBefore = Math.Max(0D, spacingBefore - pendingSpacingAfter.Value);
                }

                double spacingAfter = GetNativeCellParagraphSpacingAfter(paragraph, nativeDefaults, tableStyleDefaults);
                if (ShouldSuppressNativeContextualSpacingAfter(paragraph, GetNextNativeRenderableCellParagraph(cellParagraphs, i, footnoteNumbersById, tableStyleDefaults, nativeDefaults))) {
                    spacingAfter = 0D;
                }

                paragraphs.Add(new PdfCore.PdfTableCellParagraph(
                    paragraphRuns,
                    spacingAfter,
                    MapNativeParagraphAlign(ResolveNativeParagraphJustification(paragraph)),
                    spacingBefore));
                pendingSpacingAfter = spacingAfter;
            }

            return new NativeCellText(runs, paragraphs);
        }

        private static WordParagraph? GetNextNativeRenderableCellParagraph(IReadOnlyList<WordParagraph> paragraphs, int index, Dictionary<long, int>? footnoteNumbersById, NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults) {
            for (int nextIndex = index + 1; nextIndex < paragraphs.Count; nextIndex++) {
                WordParagraph next = paragraphs[nextIndex];
                if (CreateNativeCellParagraphRuns(next, footnoteNumbersById, tableStyleDefaults, nativeDefaults).Count > 0) {
                    return next;
                }
            }

            return null;
        }

        private static double GetNativeCellParagraphSpacingBefore(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeTableStyleDefaults tableStyleDefaults) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults);
            double lineHeight = ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, styleDefaults);
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            double spacingBefore = paragraph.LineSpacingBeforePoints ??
                GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) ??
                styleDefaults.SpacingBefore ??
                tableStyleDefaults.ParagraphSpacingBefore ??
                (nativeDefaults.ParagraphSpacingBeforeDeclared ? nativeDefaults.ParagraphSpacingBefore : 0D);
            return spacingBefore > 0D && !double.IsNaN(spacingBefore) && !double.IsInfinity(spacingBefore) ? spacingBefore : 0D;
        }

        private static double GetNativeCellParagraphSpacingAfter(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults) {
            return GetNativeCellParagraphSpacingAfter(paragraph, nativeDefaults, NativeTableStyleDefaults.Empty);
        }

        private static double GetNativeCellParagraphSpacingAfter(WordParagraph paragraph, NativeDocumentDefaults nativeDefaults, NativeTableStyleDefaults tableStyleDefaults) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            double fontSize = ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults);
            double lineHeight = ResolveNativeParagraphLineHeight(paragraph, fontSize, nativeDefaults, styleDefaults);
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            double spacingAfter = paragraph.LineSpacingAfterPoints ??
                GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) ??
                styleDefaults.SpacingAfter ??
                tableStyleDefaults.ParagraphSpacingAfter ??
                (nativeDefaults.ParagraphSpacingAfterDeclared ? nativeDefaults.ParagraphSpacingAfter : 0D);
            return spacingAfter > 0D && !double.IsNaN(spacingAfter) && !double.IsInfinity(spacingAfter) ? spacingAfter : 0D;
        }

        private static List<PdfCore.TextRun> CreateNativeCellParagraphRuns(WordParagraph paragraph, Dictionary<long, int>? footnoteNumbersById) =>
            CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById, NativeTableStyleDefaults.Empty, GetNativeDocumentDefaults(paragraph._document));

        private static List<PdfCore.TextRun> CreateNativeCellParagraphRuns(WordParagraph paragraph, Dictionary<long, int>? footnoteNumbersById, NativeTableStyleDefaults tableStyleDefaults) {
            return CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById, tableStyleDefaults, GetNativeDocumentDefaults(paragraph._document));
        }

        private static List<PdfCore.TextRun> CreateNativeCellParagraphRuns(WordParagraph paragraph, Dictionary<long, int>? footnoteNumbersById, NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults) {
            var result = new List<PdfCore.TextRun>();
            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string content = paragraph.IsHyperLink && paragraph.Hyperlink != null ? paragraph.Hyperlink.Text : AppendNativeTextWithEquation(paragraph.Text, paragraph);
            bool hasRenderableRuns = runs.Any(run => !run.IsImage && !string.IsNullOrEmpty(run.Text));
            IReadOnlyList<WordTabStop> tabStops = paragraph.TabStops;
            int tabIndex = 0;
            IReadOnlyList<W.SdtRun> repeatingSectionControls = GetNativeRepeatingSectionControls(paragraph);

            if (hasRenderableRuns) {
                foreach (WordParagraph run in runs) {
                    if (run.IsImage && run.Image != null) {
                        continue;
                    }

                    if (IsNativeTextWrappingBreak(run) && string.IsNullOrEmpty(run.Text)) {
                        result.Add(PdfCore.TextRun.LineBreak());
                        tabIndex = 0;
                        continue;
                    }

                    AddNativeCellRun(result, run, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
                }

                string? supplementalText = GetNativeSupplementalTextAfterRuns(content, runs);
                if (!string.IsNullOrEmpty(supplementalText)) {
                    AddNativeCellText(result, supplementalText!, paragraph, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
                }
            } else if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !string.IsNullOrEmpty(paragraph.Hyperlink.Text)) {
                AddNativeCellHyperLinkRun(result, paragraph.Hyperlink.Text, paragraph, paragraph.Hyperlink, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
            } else if (!string.IsNullOrEmpty(content)) {
                AddNativeCellText(result, content, paragraph, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
            }

            foreach (W.SdtRun repeatingSection in repeatingSectionControls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(repeatingSection)) {
                    if (string.IsNullOrWhiteSpace(itemText)) {
                        continue;
                    }

                    if (result.Count > 0) {
                        result.Add(PdfCore.TextRun.LineBreak());
                        tabIndex = 0;
                    }

                    AddNativeCellText(result, itemText, paragraph, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
                }
            }

            if (footnoteNumbersById != null) {
                List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, runs, Array.Empty<int>(), footnoteNumbersById);
                AddNativeCellFootnoteReferences(result, paragraphFootnoteNumbers);
            }

            return result;
        }

        private static void AddNativeCellRun(List<PdfCore.TextRun> target, WordParagraph run, NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            if (string.IsNullOrEmpty(run.Text)) {
                return;
            }

            if (run.IsHyperLink && run.Hyperlink != null) {
                AddNativeCellHyperLinkRun(target, run.Text, run, run.Hyperlink, tableStyleDefaults, nativeDefaults, tabStops, ref tabIndex);
                return;
            }

            AddNativeCellTextRuns(target, run.Text, text => CreateNativeCellTextRun(text, run, tableStyleDefaults, nativeDefaults), tabStops, ref tabIndex);
        }

        private static void AddNativeCellText(List<PdfCore.TextRun> target, string text, WordParagraph paragraph, NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            AddNativeCellTextRuns(target, text, value => CreateNativeCellTextRun(value, paragraph, tableStyleDefaults, nativeDefaults), tabStops, ref tabIndex);
        }

        private static void AddNativeCellHyperLinkRun(List<PdfCore.TextRun> target, string text, WordParagraph paragraph, WordHyperLink hyperlink, NativeTableStyleDefaults tableStyleDefaults, NativeDocumentDefaults nativeDefaults, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            AddNativeCellTextRuns(target, text, value => CreateNativeCellLinkRun(value, paragraph, hyperlink, tableStyleDefaults, nativeDefaults), tabStops, ref tabIndex);
        }

        private static void AddNativeCellTextRuns(List<PdfCore.TextRun> target, string text, Func<string, PdfCore.TextRun> createRun, IReadOnlyList<WordTabStop> tabStops, ref int tabIndex) {
            int currentTabIndex = tabIndex;
            AddNativeTextSegments(
                text,
                value => AddOrMergeNativeCellTextRun(target, createRun(value)),
                () => target.Add(PdfCore.TextRun.LineBreak()),
                () => {
                    target.Add(CreateNativeCellTabRun(tabStops, currentTabIndex));
                    currentTabIndex++;
                },
                () => currentTabIndex = 0);
            tabIndex = currentTabIndex;
        }

        private static void AddOrMergeNativeCellTextRun(List<PdfCore.TextRun> target, PdfCore.TextRun run) {
            if (target.Count == 0 || !CanMergeNativeCellTextRuns(target[target.Count - 1], run)) {
                target.Add(run);
                return;
            }

            PdfCore.TextRun previous = target[target.Count - 1];
            target[target.Count - 1] = new PdfCore.TextRun(
                previous.Text + run.Text,
                bold: previous.Bold,
                underline: previous.Underline,
                color: previous.Color,
                italic: previous.Italic,
                strike: previous.Strike,
                fontSize: previous.FontSize,
                font: previous.Font,
                baseline: previous.Baseline,
                backgroundColor: previous.BackgroundColor);
        }

        private static bool CanMergeNativeCellTextRuns(PdfCore.TextRun left, PdfCore.TextRun right) =>
            left.LinkUri == null &&
            left.LinkDestinationName == null &&
            right.LinkUri == null &&
            right.LinkDestinationName == null &&
            left.TabLeader == PdfCore.PdfTabLeaderStyle.None &&
            right.TabLeader == PdfCore.PdfTabLeaderStyle.None &&
            left.TabAlignment == PdfCore.PdfTabAlignment.Left &&
            right.TabAlignment == PdfCore.PdfTabAlignment.Left &&
            left.Text != "\n" &&
            left.Text != "\t" &&
            right.Text != "\n" &&
            right.Text != "\t" &&
            left.Bold == right.Bold &&
            left.Underline == right.Underline &&
            left.Italic == right.Italic &&
            left.Strike == right.Strike &&
            NullableDoubleEquals(left.FontSize, right.FontSize) &&
            left.Font == right.Font &&
            left.Baseline == right.Baseline &&
            Equals(left.Color, right.Color) &&
            Equals(left.BackgroundColor, right.BackgroundColor);

        private static PdfCore.TextRun CreateNativeCellTextRun(string text, WordParagraph paragraph, NativeTableStyleDefaults tableStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null) {
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, tableRunStyleDefaults: tableStyleDefaults.RunStyle, nativeDefaults: nativeDefaults);
            return new PdfCore.TextRun(
                text,
                bold: style.Bold,
                underline: style.Underline,
                color: style.Color,
                italic: style.Italic,
                strike: style.Strike,
                fontSize: style.FontSize,
                font: style.Font,
                baseline: GetNativeTextBaseline(paragraph),
                backgroundColor: style.BackgroundColor);
        }

        private static PdfCore.TextRun CreateNativeCellLinkRun(string text, WordParagraph paragraph, WordHyperLink hyperlink, NativeTableStyleDefaults tableStyleDefaults = default, NativeDocumentDefaults? nativeDefaults = null) {
            Uri? uri = hyperlink.Uri;
            string? linkUri = uri != null && uri.IsAbsoluteUri ? uri.AbsoluteUri : null;
            string? destinationName = linkUri != null || string.IsNullOrWhiteSpace(hyperlink.Anchor) ? null : hyperlink.Anchor;
            if (linkUri == null && destinationName == null) {
                return CreateNativeCellTextRun(text, paragraph, tableStyleDefaults, nativeDefaults);
            }

            string? contents = string.IsNullOrWhiteSpace(hyperlink.Tooltip) ? null : hyperlink.Tooltip;
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(paragraph, tableRunStyleDefaults: tableStyleDefaults.RunStyle, nativeDefaults: nativeDefaults);
            return new PdfCore.TextRun(
                text,
                bold: style.Bold,
                underline: style.Underline || linkUri != null || destinationName != null,
                color: style.Color,
                italic: style.Italic,
                strike: style.Strike,
                fontSize: style.FontSize,
                font: style.Font,
                linkUri: linkUri,
                linkContents: contents,
                baseline: GetNativeTextBaseline(paragraph),
                linkDestinationName: destinationName,
                backgroundColor: style.BackgroundColor);
        }

        private static PdfCore.TextRun CreateNativeCellTabRun(IReadOnlyList<WordTabStop> tabStops, int tabIndex) {
            if (tabIndex < tabStops.Count) {
                WordTabStop tabStop = tabStops[tabIndex];
                return PdfCore.TextRun.Tab(MapNativeTabLeader(tabStop.Leader), MapNativeTabAlignment(tabStop.Alignment));
            }

            return PdfCore.TextRun.Tab();
        }

        private static PdfCore.PdfTextBaseline GetNativeTextBaseline(WordParagraph paragraph) =>
            paragraph.VerticalTextAlignment == W.VerticalPositionValues.Superscript
                ? PdfCore.PdfTextBaseline.Superscript
                : paragraph.VerticalTextAlignment == W.VerticalPositionValues.Subscript
                    ? PdfCore.PdfTextBaseline.Subscript
                    : PdfCore.PdfTextBaseline.Normal;

        private static void AddNativeCellFootnoteReferences(List<PdfCore.TextRun> target, IReadOnlyList<int> footnoteNumbers) {
            foreach (int footnoteNumber in footnoteNumbers) {
                target.Add(PdfCore.TextRun.Superscript(footnoteNumber.ToString(CultureInfo.InvariantCulture)));
            }
        }

        private static string GetNativeCellText(WordTableCell cell, Dictionary<long, int>? footnoteNumbersById) {
            var parts = new List<string>();
            foreach (WordParagraph paragraph in GetNativeCellParagraphs(cell)) {
                string? paragraphText = GetNativeCellParagraphText(paragraph);
                if (!string.IsNullOrEmpty(paragraphText)) {
                    string text = paragraphText;
                    if (footnoteNumbersById != null) {
                        List<int> paragraphFootnoteNumbers = GetNativeParagraphFootnoteNumbers(paragraph, GetNativeRuns(paragraph), Array.Empty<int>(), footnoteNumbersById);
                        if (paragraphFootnoteNumbers.Count > 0) {
                            text += string.Concat(paragraphFootnoteNumbers.Select(number => number.ToString(CultureInfo.InvariantCulture)));
                        }
                    }

                    parts.Add(text);
                }
            }

            return string.Join(Environment.NewLine, parts);
        }

        private static IReadOnlyList<WordParagraph> GetNativeCellParagraphs(WordTableCell cell) =>
            CollapseNativeParagraphElements(cell.Paragraphs.Cast<WordElement>())
                .OfType<WordParagraph>()
                .ToList();

        private static string GetNativeCellParagraphText(WordParagraph paragraph) {
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !string.IsNullOrEmpty(paragraph.Hyperlink.Text)) {
                return paragraph.Hyperlink.Text;
            }

            if (!string.IsNullOrEmpty(paragraph.Text)) {
                return AppendNativeTextWithEquation(paragraph.Text, paragraph);
            }

            var parts = new List<string>();
            foreach (WordParagraph run in paragraph.GetRuns()) {
                string runText = run.IsHyperLink && run.Hyperlink != null ? run.Hyperlink.Text : run.Text;
                if (!string.IsNullOrEmpty(runText)) {
                    parts.Add(runText);
                }
            }

            string text = string.Concat(parts);
            return AppendNativeTextWithEquation(text, paragraph);
        }

    }
}
