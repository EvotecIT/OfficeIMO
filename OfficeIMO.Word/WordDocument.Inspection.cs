#pragma warning disable CS1591

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        public WordDocumentSnapshot CreateInspectionSnapshot() {
            var snapshot = new WordDocumentSnapshot {
                FilePath = string.IsNullOrWhiteSpace(FilePath) ? null : FilePath,
                Title = BuiltinDocumentProperties?.Title,
            };

            for (int sectionIndex = 0; sectionIndex < Sections.Count; sectionIndex++) {
                var section = Sections[sectionIndex];
                var sectionSnapshot = new WordSectionSnapshot {
                    Index = sectionIndex,
                    SectionBreakType = ResolveSectionBreakType(section, sectionIndex),
                    Orientation = NormalizeOpenXmlEnumValue(section._sectionProperties?.GetFirstChild<PageSize>()?.Orient) ?? section.PageOrientation.ToString(),
                    PageWidthPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageSize>()?.Width?.Value),
                    PageHeightPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageSize>()?.Height?.Value),
                    MarginTopPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Top?.Value),
                    MarginBottomPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Bottom?.Value),
                    MarginLeftPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Left?.Value),
                    MarginRightPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Right?.Value),
                    HeaderMarginPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Header?.Value),
                    FooterMarginPoints = ConvertTwipsToPoints(section._sectionProperties?.GetFirstChild<PageMargin>()?.Footer?.Value),
                    ColumnCount = section.ColumnCount,
                    ColumnSpacingPoints = ConvertTwipsToPoints(section.ColumnsSpace),
                    HasColumnSeparator = section._sectionProperties?.GetFirstChild<Columns>()?.Separator?.Value ?? false,
                    PageNumberStart = section._sectionProperties?.GetFirstChild<PageNumberType>()?.Start?.Value,
                    HeaderCount = CountHeaderParts(section.Header),
                    FooterCount = CountFooterParts(section.Footer),
                    DifferentFirstPage = section.DifferentFirstPage,
                    DifferentOddAndEvenPages = section.DifferentOddAndEvenPages,
                    DefaultHeader = BuildHeaderFooterSnapshot(section.Header?.Default, "header", "default"),
                    DefaultFooter = BuildHeaderFooterSnapshot(section.Footer?.Default, "footer", "default"),
                    FirstHeader = BuildHeaderFooterSnapshot(section.Header?.First, "header", "first"),
                    FirstFooter = BuildHeaderFooterSnapshot(section.Footer?.First, "footer", "first"),
                    EvenHeader = BuildHeaderFooterSnapshot(section.Header?.Even, "header", "even"),
                    EvenFooter = BuildHeaderFooterSnapshot(section.Footer?.Even, "footer", "even"),
                };

                int order = 0;
                var elements = section.Elements;
                for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
                    var element = elements[elementIndex];
                    if (element is WordParagraph paragraphPart) {
                        var group = new List<WordParagraph> { paragraphPart };
                        while (elementIndex + 1 < elements.Count
                               && elements[elementIndex + 1] is WordParagraph nextParagraph
                               && ReferenceEquals(nextParagraph._paragraph, paragraphPart._paragraph)) {
                            group.Add(nextParagraph);
                            elementIndex++;
                        }

                        var paragraphSnapshot = BuildParagraphSnapshot(new WordParagraph(this, paragraphPart._paragraph));
                        paragraphSnapshot.Order = order++;
                        sectionSnapshot.AddElement(paragraphSnapshot);
                    } else if (element is WordTable table) {
                        var tableSnapshot = BuildTableSnapshot(table);
                        tableSnapshot.Order = order++;
                        sectionSnapshot.AddElement(tableSnapshot);
                    }
                }

                snapshot.AddSection(sectionSnapshot);
            }

            return snapshot;
        }

        private static int CountHeaderParts(WordHeaders? collection) {
            if (collection == null) {
                return 0;
            }

            int count = 0;
            if (collection.Default != null) count++;
            if (collection.Even != null) count++;
            if (collection.First != null) count++;
            return count;
        }

        private static int CountFooterParts(WordFooters? collection) {
            if (collection == null) {
                return 0;
            }

            int count = 0;
            if (collection.Default != null) count++;
            if (collection.Even != null) count++;
            if (collection.First != null) count++;
            return count;
        }

        private static string? ResolveSectionBreakType(WordSection section, int sectionIndex) {
            if (section == null || sectionIndex <= 0) {
                return null;
            }

            var sectionType = section._sectionProperties?.GetFirstChild<SectionType>()?.Val;
            var normalizedType = NormalizeOpenXmlEnumValue(sectionType);
            return string.IsNullOrWhiteSpace(normalizedType) ? "NextPage" : normalizedType;
        }

        private WordHeaderFooterSnapshot? BuildHeaderFooterSnapshot(
            WordHeaderFooter? headerFooter,
            string kind,
            string variant) {
            if (headerFooter == null) {
                return null;
            }

            var snapshot = new WordHeaderFooterSnapshot {
                Kind = kind,
                Variant = variant,
                TableCount = headerFooter.Tables.Count,
            };

            int order = 0;
            foreach (var child in EnumerateHeaderFooterBlocks(headerFooter)) {
                switch (child) {
                    case Paragraph paragraph:
                        var paragraphSnapshot = BuildParagraphSnapshot(new WordParagraph(this, paragraph));
                        paragraphSnapshot.Order = order++;
                        snapshot.AddElement(paragraphSnapshot);
                        break;
                    case Table table:
                        var tableSnapshot = BuildTableSnapshot(new WordTable(this, table));
                        tableSnapshot.Order = order++;
                        snapshot.AddElement(tableSnapshot);
                        break;
                }
            }

            return snapshot;
        }

        private WordParagraphSnapshot BuildParagraphSnapshot(WordParagraph paragraph) {
            var bookmark = paragraph.Bookmark;
            var bookmarkStart = paragraph._paragraph.ChildElements.OfType<BookmarkStart>().FirstOrDefault();
            var snapshot = new WordParagraphSnapshot {
                Text = string.Concat(paragraph.GetRuns().Select(run => run.Text)),
                StyleId = paragraph.StyleId,
                StyleName = paragraph.Style?.ToString(),
                IsListItem = paragraph.IsListItem,
                IsOrderedList = ResolveOrderedList(paragraph),
                ListLevel = paragraph.ListItemLevel,
                ListStyleName = paragraph.ListStyle?.ToString(),
                Alignment = NormalizeOpenXmlEnumValue(paragraph._paragraphProperties?.Justification?.Val),
                IndentStartPoints = paragraph.IndentationBeforePoints,
                IndentEndPoints = paragraph.IndentationAfterPoints,
                IndentFirstLinePoints = ResolveIndentFirstLinePoints(paragraph),
                SpaceAbovePoints = paragraph.LineSpacingBeforePoints,
                SpaceBelowPoints = paragraph.LineSpacingAfterPoints,
                LineSpacingValue = paragraph.LineSpacing,
                LineSpacingRule = NormalizeOpenXmlEnumValue(paragraph.LineSpacingRule),
                ShadingFillColorHex = NormalizeColorHex(paragraph.ShadingFillColorHex),
                LeftBorder = BuildParagraphBorderSnapshot(
                    NormalizeOpenXmlEnumValue(paragraph.Borders.LeftStyle),
                    NormalizeColorHex(paragraph.Borders.LeftColorHex),
                    paragraph.Borders.LeftSize?.Value,
                    paragraph.Borders.LeftSpace?.Value),
                RightBorder = BuildParagraphBorderSnapshot(
                    NormalizeOpenXmlEnumValue(paragraph.Borders.RightStyle),
                    NormalizeColorHex(paragraph.Borders.RightColorHex),
                    paragraph.Borders.RightSize?.Value,
                    paragraph.Borders.RightSpace?.Value),
                TopBorder = BuildParagraphBorderSnapshot(
                    NormalizeOpenXmlEnumValue(paragraph.Borders.TopStyle),
                    NormalizeColorHex(paragraph.Borders.TopColorHex),
                    paragraph.Borders.TopSize?.Value,
                    paragraph.Borders.TopSpace?.Value),
                BottomBorder = BuildParagraphBorderSnapshot(
                    NormalizeOpenXmlEnumValue(paragraph.Borders.BottomStyle),
                    NormalizeColorHex(paragraph.Borders.BottomColorHex),
                    paragraph.Borders.BottomSize?.Value,
                    paragraph.Borders.BottomSpace?.Value),
                IsRightToLeft = paragraph.BiDi,
                KeepWithNext = paragraph.KeepWithNext,
                KeepLinesTogether = paragraph.KeepLinesTogether,
                AvoidWidowAndOrphan = paragraph.AvoidWidowAndOrphan,
                PageBreakBefore = paragraph.PageBreakBefore,
                BookmarkName = bookmark?.Name ?? bookmarkStart?.Name,
                BookmarkId = bookmark != null
                    ? bookmark.Id
                    : int.TryParse(bookmarkStart?.Id?.Value, out var bookmarkId) ? bookmarkId : null,
            };

            foreach (var run in paragraph.GetRuns()) {
                var hyperlink = run.Hyperlink;
                var image = run.Image;

                snapshot.AddRun(new WordRunSnapshot {
                    Text = run.Text,
                    Bold = run.Bold,
                    Italic = run.Italic,
                    Underline = run.Underline != null,
                    Strike = run.Strike || run.DoubleStrike,
                    FontSize = run.FontSize,
                    FontFamily = run.FontFamily,
                    ColorHex = NormalizeColorHex(run.ColorHex),
                    HighlightColor = NormalizeOpenXmlEnumValue(run.Highlight),
                    VerticalTextAlignment = NormalizeOpenXmlEnumValue(run.VerticalTextAlignment),
                    CapsStyle = run.CapsStyle == CapsStyle.None ? null : run.CapsStyle.ToString(),
                    IsHyperlink = hyperlink != null,
                    HyperlinkUri = hyperlink?.Uri?.ToString(),
                    HyperlinkAnchor = hyperlink?.Anchor,
                    Footnote = BuildFootnoteSnapshot(run.FootNote),
                    InlineImage = image == null ? null : new WordInlineImageSnapshot {
                        FilePath = string.IsNullOrWhiteSpace(image.FilePath) ? null : image.FilePath,
                        FileName = image.FileName,
                        ContentType = ResolveImageContentType(image),
                        Bytes = image.IsExternal ? null : image.GetBytes(),
                        Description = image.Description,
                        Title = image.Title,
                        Width = image.Width,
                        Height = image.Height,
                        IsInline = image.WrapText == WrapTextImage.InLineWithText,
                        WrapText = image.WrapText?.ToString(),
                    },
                });
            }

            foreach (var tabStop in paragraph.TabStops) {
                snapshot.AddTabStop(new WordTabStopSnapshot {
                    Alignment = NormalizeOpenXmlEnumValue(tabStop.Alignment),
                    Leader = NormalizeOpenXmlEnumValue(tabStop.Leader),
                    PositionPoints = Helpers.ConvertTwipsToPoints(tabStop.Position),
                });
            }

            return snapshot;
        }

        private WordTableSnapshot BuildTableSnapshot(WordTable table) {
            var snapshot = new WordTableSnapshot {
                RowCount = table.Rows.Count,
                ColumnCount = table.Rows.Count == 0 ? 0 : table.Rows.Max(row => row.CellsCount),
                StyleName = table.Style?.ToString(),
                Title = table.Title,
                Description = table.Description,
                RepeatHeaderRow = table.Rows.Count > 0 && table.RepeatAsHeaderRowAtTheTopOfEachPage,
            };

            var gridColumnWidths = table.GridColumnWidth;
            if (gridColumnWidths.Count > 0) {
                foreach (var width in gridColumnWidths) {
                    snapshot.AddColumnWidth(Helpers.ConvertTwipsToPoints(width));
                }
            } else {
                foreach (var width in table.ColumnWidth) {
                    snapshot.AddColumnWidth(Helpers.ConvertTwipsToPoints(width));
                }
            }

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                var row = table.Rows[rowIndex];
                var rowSnapshot = new WordTableRowSnapshot {
                    RowIndex = rowIndex,
                };

                for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++) {
                    var cell = row.Cells[columnIndex];
                    var cellSnapshot = new WordTableCellSnapshot {
                        ColumnIndex = columnIndex,
                        ColumnSpan = ResolveColumnSpan(cell, row, columnIndex),
                        RowSpan = ResolveRowSpan(table, rowIndex, columnIndex),
                        ShadingFillColorHex = NormalizeColorHex(cell.ShadingFillColorHex),
                        LeftBorder = BuildBorderSnapshot(
                            NormalizeOpenXmlEnumValue(cell.Borders.LeftStyle),
                            NormalizeColorHex(cell.Borders.LeftColorHex),
                            cell.Borders.LeftSize?.Value),
                        RightBorder = BuildBorderSnapshot(
                            NormalizeOpenXmlEnumValue(cell.Borders.RightStyle),
                            NormalizeColorHex(cell.Borders.RightColorHex),
                            cell.Borders.RightSize?.Value),
                        TopBorder = BuildBorderSnapshot(
                            NormalizeOpenXmlEnumValue(cell.Borders.TopStyle),
                            NormalizeColorHex(cell.Borders.TopColorHex),
                            cell.Borders.TopSize?.Value),
                        BottomBorder = BuildBorderSnapshot(
                            NormalizeOpenXmlEnumValue(cell.Borders.BottomStyle),
                            NormalizeColorHex(cell.Borders.BottomColorHex),
                            cell.Borders.BottomSize?.Value),
                        HasHorizontalMerge = cell.HasHorizontalMerge,
                        HasVerticalMerge = cell.HasVerticalMerge,
                    };

                    snapshot.HasHorizontalMerges |= cell.HasHorizontalMerge;
                    snapshot.HasVerticalMerges |= cell.HasVerticalMerge;

                    foreach (var paragraphGroup in GroupParagraphs(cell.Paragraphs)) {
                        cellSnapshot.AddParagraph(BuildParagraphSnapshot(paragraphGroup));
                    }

                    rowSnapshot.AddCell(cellSnapshot);
                }

                snapshot.AddRow(rowSnapshot);
            }

            return snapshot;
        }

        private WordFootnoteSnapshot? BuildFootnoteSnapshot(WordFootNote? footNote) {
            if (footNote == null) {
                return null;
            }

            var paragraphs = footNote.Paragraphs;
            if (paragraphs == null || paragraphs.Count == 0) {
                return null;
            }

            var snapshot = new WordFootnoteSnapshot {
                ReferenceId = footNote.ReferenceId,
            };

            foreach (var paragraphGroup in GroupParagraphs(paragraphs).Where(paragraph => paragraph.GetRuns().Any(run => run.FootNote == null))) {
                snapshot.AddParagraph(BuildParagraphSnapshot(paragraphGroup));
            }

            return snapshot.Paragraphs.Count > 0 ? snapshot : null;
        }

        private static double? ResolveIndentFirstLinePoints(WordParagraph paragraph) {
            if (paragraph.IndentationFirstLinePoints.HasValue) {
                return paragraph.IndentationFirstLinePoints.Value;
            }

            if (paragraph.IndentationHangingPoints.HasValue) {
                return -paragraph.IndentationHangingPoints.Value;
            }

            return null;
        }

        private IEnumerable<WordParagraph> GroupParagraphs(IEnumerable<WordParagraph> paragraphParts) {
            Paragraph? currentParagraph = null;

            foreach (var paragraphPart in paragraphParts) {
                if (!ReferenceEquals(currentParagraph, paragraphPart._paragraph)) {
                    currentParagraph = paragraphPart._paragraph;
                    yield return new WordParagraph(this, paragraphPart._paragraph);
                }
            }
        }

        private IEnumerable<OpenXmlElement> EnumerateHeaderFooterBlocks(WordHeaderFooter headerFooter) {
            if (headerFooter._header != null) {
                foreach (var child in headerFooter._header.ChildElements) {
                    if (child is Paragraph || child is Table) {
                        yield return child;
                    }
                }

                yield break;
            }

            if (headerFooter._footer != null) {
                foreach (var child in headerFooter._footer.ChildElements) {
                    if (child is Paragraph || child is Table) {
                        yield return child;
                    }
                }
            }
        }

        private static string? NormalizeColorHex(string? value) {
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private static WordTableCellBorderSnapshot? BuildBorderSnapshot(
            string? style,
            string? colorHex,
            uint? size) {
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(colorHex) && !size.HasValue) {
                return null;
            }

            return new WordTableCellBorderSnapshot {
                Style = style,
                ColorHex = colorHex,
                Size = size,
            };
        }

        private static WordParagraphBorderSnapshot? BuildParagraphBorderSnapshot(
            string? style,
            string? colorHex,
            uint? size,
            uint? space) {
            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(colorHex) && !size.HasValue && !space.HasValue) {
                return null;
            }

            return new WordParagraphBorderSnapshot {
                Style = style,
                ColorHex = colorHex,
                Size = size,
                Space = space,
            };
        }

        private static string? ResolveImageContentType(WordImage image) {
            if (image == null) {
                return null;
            }

            var fileName = image.FileName ?? image.FilePath;
            if (string.IsNullOrWhiteSpace(fileName)) {
                return null;
            }

            switch (System.IO.Path.GetExtension(fileName).ToLowerInvariant()) {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".svg":
                    return "image/svg+xml";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                default:
                    return null;
            }
        }

        private static bool? ResolveOrderedList(WordParagraph paragraph) {
            if (paragraph == null || !paragraph.IsListItem) {
                return null;
            }

            return paragraph.ListStyle switch {
                WordListStyle.Bulleted => false,
                WordListStyle.BulletedChars => false,
                WordListStyle.Custom => null,
                null => null,
                _ => true,
            };
        }

        private static int ResolveColumnSpan(WordTableCell cell, WordTableRow row, int columnIndex) {
            var gridSpan = cell._tableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value;
            if (gridSpan.HasValue && gridSpan.Value > 1) {
                return gridSpan.Value;
            }

            if (cell.HorizontalMerge == MergedCellValues.Restart) {
                int span = 1;
                for (int index = columnIndex + 1; index < row.Cells.Count; index++) {
                    if (row.Cells[index].HorizontalMerge == MergedCellValues.Continue) {
                        span++;
                        continue;
                    }

                    break;
                }

                return span;
            }

            return 1;
        }

        private static int ResolveRowSpan(WordTable table, int rowIndex, int columnIndex) {
            if (rowIndex < 0 || rowIndex >= table.Rows.Count) {
                return 1;
            }

            if (columnIndex < 0 || columnIndex >= table.Rows[rowIndex].Cells.Count) {
                return 1;
            }

            var cell = table.Rows[rowIndex].Cells[columnIndex];
            if (cell.VerticalMerge != MergedCellValues.Restart) {
                return 1;
            }

            int span = 1;
            for (int index = rowIndex + 1; index < table.Rows.Count; index++) {
                if (columnIndex >= table.Rows[index].Cells.Count) {
                    break;
                }

                if (table.Rows[index].Cells[columnIndex].VerticalMerge == MergedCellValues.Continue) {
                    span++;
                    continue;
                }

                break;
            }

            return span;
        }

        private static string? NormalizeOpenXmlEnumValue(object? value) {
            if (value == null) {
                return null;
            }

            var innerTextProperty = value.GetType().GetProperty("InnerText");
            string? innerText = innerTextProperty?.GetValue(value, null)?.ToString();
            if (!string.IsNullOrWhiteSpace(innerText)) {
                string normalizedInnerText = innerText!;
                return char.ToUpperInvariant(normalizedInnerText[0]) + normalizedInnerText.Substring(1);
            }

            var valueField = value.GetType().GetField("_value", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            string? fieldValue = valueField?.GetValue(value)?.ToString();
            if (!string.IsNullOrWhiteSpace(fieldValue)) {
                string normalizedFieldValue = fieldValue!;
                return char.ToUpperInvariant(normalizedFieldValue[0]) + normalizedFieldValue.Substring(1);
            }

            var text = value.ToString();
            if (!string.IsNullOrWhiteSpace(text) && !text.Contains("{", StringComparison.Ordinal)) {
                return text;
            }

            var property = value.GetType().GetProperty("Value");
            object? rawValueObject = property?.GetValue(value, null);
            string? rawValue = rawValueObject as string ?? rawValueObject?.ToString();
            if (string.IsNullOrWhiteSpace(rawValue)) {
                return string.IsNullOrWhiteSpace(text) ? null : text;
            }

            string normalized = rawValue ?? string.Empty;
            return char.ToUpperInvariant(normalized[0]) + normalized.Substring(1);
        }

        private static double? ConvertTwipsToPoints(uint? twips) {
            if (!twips.HasValue) {
                return null;
            }

            return Helpers.ConvertTwipsToPoints((int)twips.Value);
        }

        private static double? ConvertTwipsToPoints(int? twips) {
            if (!twips.HasValue) {
                return null;
            }

            return Helpers.ConvertTwipsToPoints(twips.Value);
        }
    }
}
