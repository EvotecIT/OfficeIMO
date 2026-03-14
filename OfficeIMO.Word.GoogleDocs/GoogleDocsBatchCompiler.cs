using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using System.IO;

namespace OfficeIMO.Word.GoogleDocs {
    internal static class GoogleDocsBatchCompiler {
        internal static GoogleDocsBatch Build(WordDocument document, GoogleDocsSaveOptions options) {
            var plan = GoogleDocsPlanBuilder.Build(document, options);
            var report = plan.Report;
            var snapshot = document.CreateInspectionSnapshot();
            var title = ResolveTitle(document, snapshot, options);
            var batch = new GoogleDocsBatch(title, plan, report, snapshot);

            bool sectionsNoticeAdded = false;
            bool listNoticeAdded = false;
            bool styleNoticeAdded = false;
            bool imageNoticeAdded = false;
            bool footnoteNoticeAdded = false;
            bool tableNoticeAdded = false;
            bool tableMergeNoticeAdded = false;
            bool bookmarkNoticeAdded = false;
            bool capsNoticeAdded = false;
            bool lineSpacingNoticeAdded = false;
            bool tabStopNoticeAdded = false;
            bool sectionLayoutNoticeAdded = false;

            for (int sectionIndex = 0; sectionIndex < snapshot.Sections.Count; sectionIndex++) {
                var section = snapshot.Sections[sectionIndex];

                if (!sectionsNoticeAdded && snapshot.Sections.Count > 1) {
                    report.Add(
                        TranslationSeverity.Info,
                        "Sections",
                        "Multiple Word sections are preserved as distinct scopes in the neutral Google Docs batch, and paragraph-led sections can now be translated into native Google Docs section breaks.");
                    sectionsNoticeAdded = true;
                }

                if ((section.HeaderCount > 0 || section.FooterCount > 0 || section.DifferentFirstPage || section.DifferentOddAndEvenPages)
                    && section.Index > 0) {
                    report.Add(
                        TranslationSeverity.Info,
                        "HeadersAndFooters",
                        "Word section header/footer metadata is preserved in the snapshot. Default, first-page, and even-page header/footer execution are implemented through the current Google Docs export path.");
                }

                AddSegment(batch, report, section.Index, section.DefaultHeader);
                AddSegment(batch, report, section.Index, section.DefaultFooter);
                AddSegment(batch, report, section.Index, section.FirstHeader);
                AddSegment(batch, report, section.Index, section.FirstFooter);
                AddSegment(batch, report, section.Index, section.EvenHeader);
                AddSegment(batch, report, section.Index, section.EvenFooter);

                for (int elementIndex = 0; elementIndex < section.Elements.Count; elementIndex++) {
                    var element = section.Elements[elementIndex];
                    if (element is WordParagraphSnapshot paragraph) {
                        bool startsNewSectionBefore = section.Index > 0
                            && elementIndex == 0
                            && !string.IsNullOrWhiteSpace(section.SectionBreakType);

                        if (!listNoticeAdded && paragraph.IsListItem) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Lists",
                                "Word list semantics now flow into the neutral Google Docs batch with ordered-vs-bulleted classification, which gives the execution layer enough information to emit native Google Docs list requests.");
                            listNoticeAdded = true;
                        }

                        if (section.Index > 0
                            && elementIndex == 0
                            && !string.IsNullOrWhiteSpace(section.SectionBreakType)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "SectionBreaks",
                                "Word section breaks that start with a paragraph can now be emitted as native Google Docs insertSectionBreak requests.");
                        }

                        if (!styleNoticeAdded && HasFormatting(paragraph)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "TextStyles",
                                "Paragraph and run formatting are preserved in the neutral Google Docs batch so a later execution layer can translate them to updateParagraphStyle and updateTextStyle requests.");
                            styleNoticeAdded = true;
                        }

                        if (!capsNoticeAdded && paragraph.Runs.Any(run => string.Equals(run.CapsStyle, nameof(CapsStyle.Caps), StringComparison.OrdinalIgnoreCase))) {
                            report.Add(
                                TranslationSeverity.Warning,
                                "TextStyles",
                                "Word all-caps run styling is preserved in the neutral Google Docs batch, but the current Google Docs export only emits native smallCaps. Full caps remains a dedicated follow-up slice.");
                            capsNoticeAdded = true;
                        }

                        if (!lineSpacingNoticeAdded && paragraph.LineSpacingValue.HasValue && !CanMapLineSpacing(paragraph)) {
                            report.Add(
                                TranslationSeverity.Warning,
                                "ParagraphStyles",
                                "Some Word line spacing rules still do not have a dependable Google Docs approximation. Unsupported rules remain a dedicated follow-up slice.");
                            lineSpacingNoticeAdded = true;
                        }

                        if (paragraph.LineSpacingValue.HasValue
                            && CanMapLineSpacing(paragraph)
                            && !string.Equals(paragraph.LineSpacingRule, "Auto", StringComparison.OrdinalIgnoreCase)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "ParagraphStyles",
                                "Word Exact and AtLeast line spacing are now exported as Google Docs lineSpacing approximations based on the Word line-height value.");
                        }

                        if (!tabStopNoticeAdded && paragraph.TabStops.Count > 0) {
                            report.Add(
                                paragraph.TabStops.Any(tabStop => !string.IsNullOrWhiteSpace(tabStop.Leader) && !string.Equals(tabStop.Leader, "None", StringComparison.OrdinalIgnoreCase))
                                    ? TranslationSeverity.Warning
                                    : TranslationSeverity.Info,
                                "ParagraphStyles",
                                paragraph.TabStops.Any(tabStop => !string.IsNullOrWhiteSpace(tabStop.Leader) && !string.Equals(tabStop.Leader, "None", StringComparison.OrdinalIgnoreCase))
                                    ? "Word tab stops are now emitted as native Google Docs paragraph tabStops, but Word tab leaders do not have a direct Google Docs paragraph-style equivalent in the current export path."
                                    : "Word tab stops are now emitted as native Google Docs paragraph tabStops.");
                            tabStopNoticeAdded = true;
                        }

                        if (!imageNoticeAdded && paragraph.Runs.Any(run => run.InlineImage != null)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "InlineImages",
                                "Inline Word images are preserved in the neutral Google Docs batch and can now be materialized as native Google Docs inline images through the exporter pipeline.");
                            imageNoticeAdded = true;
                        }

                        if (!footnoteNoticeAdded && paragraph.Runs.Any(run => run.Footnote != null)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Footnotes",
                                "Word footnotes are preserved in the neutral Google Docs batch and can now be emitted as native Google Docs footnotes for body paragraphs.");
                            footnoteNoticeAdded = true;
                        }

                        if (!bookmarkNoticeAdded && (!string.IsNullOrWhiteSpace(paragraph.BookmarkName) || paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.HyperlinkAnchor)))) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Bookmarks",
                                "Word paragraph bookmarks are preserved in the neutral Google Docs batch and can now be emitted as native Google Docs named ranges in body flow, header/footer segments, footnotes, and replayed table cells. Internal anchor links remain preserved for a dedicated follow-up linking slice.");
                            bookmarkNoticeAdded = true;
                        }

                        batch.Add(new GoogleDocsInsertParagraphRequest {
                            SectionIndex = sectionIndex,
                            ElementIndex = elementIndex,
                            SectionStyle = elementIndex == 0 ? ConvertSectionStyle(section) : null,
                            Paragraph = ConvertParagraph(paragraph, startsNewSectionBefore, section.SectionBreakType),
                        });

                        if (!sectionLayoutNoticeAdded && elementIndex == 0 && HasSectionLayout(section)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Sections",
                                "Word section page size and supported margins are now preserved in the neutral Google Docs batch so the exporter can emit native updateSectionStyle requests.");
                            sectionLayoutNoticeAdded = true;
                        }
                    } else if (element is WordTableSnapshot table) {
                        bool startsNewSectionBefore = section.Index > 0
                            && elementIndex == 0
                            && !string.IsNullOrWhiteSpace(section.SectionBreakType);

                        if (startsNewSectionBefore) {
                            report.Add(
                                TranslationSeverity.Info,
                                "SectionBreaks",
                                "Word section breaks that start with a table can now be emitted as native Google Docs insertSectionBreak requests before the leading table block.");
                        }

                        if (!tableNoticeAdded) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Tables",
                                "Word table structure is now compiled into a neutral Google Docs table model, which gives the export path a real foundation beyond planning-only analysis.");
                            tableNoticeAdded = true;
                        }

                        if (!tableMergeNoticeAdded && (table.HasHorizontalMerges || table.HasVerticalMerges)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "TableMerges",
                                "Merged Word table cells are preserved in the neutral Google Docs batch and can now be replayed through native Google Docs mergeTableCells requests.");
                            tableMergeNoticeAdded = true;
                        }

                        batch.Add(new GoogleDocsInsertTableRequest {
                            SectionIndex = sectionIndex,
                            ElementIndex = elementIndex,
                            SectionStyle = elementIndex == 0 ? ConvertSectionStyle(section) : null,
                            Table = ConvertTable(table),
                            StartsNewSectionBefore = startsNewSectionBefore,
                            SectionBreakType = section.SectionBreakType,
                        });

                        if (!sectionLayoutNoticeAdded && elementIndex == 0 && HasSectionLayout(section)) {
                            report.Add(
                                TranslationSeverity.Info,
                                "Sections",
                                "Word section page size and supported margins are now preserved in the neutral Google Docs batch so the exporter can emit native updateSectionStyle requests.");
                            sectionLayoutNoticeAdded = true;
                        }
                    }
                }
            }

            return batch;
        }

        private static void AddSegment(
            GoogleDocsBatch batch,
            TranslationReport report,
            int sectionIndex,
            WordHeaderFooterSnapshot? source) {
            if (source == null || source.Elements.Count == 0) {
                return;
            }

            var segment = new GoogleDocsSegment {
                SectionIndex = sectionIndex,
                Kind = source.Kind,
                Variant = source.Variant,
                TableCount = source.TableCount,
            };

            foreach (var element in source.Elements) {
                switch (element) {
                    case WordParagraphSnapshot paragraph:
                        segment.AddRequest(new GoogleDocsInsertParagraphRequest {
                            SectionIndex = sectionIndex,
                            ElementIndex = element.Order,
                            Paragraph = ConvertParagraph(paragraph),
                        });
                        break;
                    case WordTableSnapshot table:
                        segment.AddRequest(new GoogleDocsInsertTableRequest {
                            SectionIndex = sectionIndex,
                            ElementIndex = element.Order,
                            Table = ConvertTable(table),
                        });
                        break;
                }
            }

            if (source.TableCount > 0) {
                report.Add(
                    TranslationSeverity.Info,
                    "HeadersAndFooters",
                    $"A {source.Kind} {source.Variant} segment contains {source.TableCount} table block(s); the Google Docs exporter now replays simple header/footer tables using the same staged insertion and cell-fill path as body tables.");
            }

            batch.AddSegment(segment);
        }

        private static string ResolveTitle(WordDocument document, WordDocumentSnapshot snapshot, GoogleDocsSaveOptions options) {
            if (!string.IsNullOrWhiteSpace(options.Title)) {
                return options.Title!;
            }

            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            if (!string.IsNullOrWhiteSpace(snapshot.FilePath)) {
                return Path.GetFileNameWithoutExtension(snapshot.FilePath);
            }

            if (!string.IsNullOrWhiteSpace(document.FilePath)) {
                return Path.GetFileNameWithoutExtension(document.FilePath);
            }

            return "Document";
        }

        private static bool HasFormatting(WordParagraphSnapshot paragraph) {
            return !string.IsNullOrWhiteSpace(paragraph.StyleId)
                || !string.IsNullOrWhiteSpace(paragraph.Alignment)
                || paragraph.IndentStartPoints.HasValue
                || paragraph.IndentEndPoints.HasValue
                || paragraph.IndentFirstLinePoints.HasValue
                || paragraph.SpaceAbovePoints.HasValue
                || paragraph.SpaceBelowPoints.HasValue
                || paragraph.LineSpacingValue.HasValue
                || !string.IsNullOrWhiteSpace(paragraph.ShadingFillColorHex)
                || paragraph.LeftBorder != null
                || paragraph.RightBorder != null
                || paragraph.TopBorder != null
                || paragraph.BottomBorder != null
                || paragraph.IsRightToLeft
                || paragraph.KeepWithNext
                || paragraph.KeepLinesTogether
                || paragraph.AvoidWidowAndOrphan
                || paragraph.PageBreakBefore
                || paragraph.TabStops.Count > 0
                || paragraph.Runs.Any(run =>
                    run.Bold
                    || run.Italic
                    || run.Underline
                    || run.Strike
                    || run.FontSize.HasValue
                    || !string.IsNullOrWhiteSpace(run.FontFamily)
                    || !string.IsNullOrWhiteSpace(run.ColorHex));
        }

        private static bool CanMapLineSpacing(WordParagraphSnapshot paragraph) {
            return ResolveLineSpacingPercent(paragraph).HasValue;
        }

        private static double? ResolveLineSpacingPercent(WordParagraphSnapshot paragraph) {
            if (!paragraph.LineSpacingValue.HasValue || paragraph.LineSpacingValue.Value <= 0) {
                return null;
            }

            if (string.Equals(paragraph.LineSpacingRule, "Auto", StringComparison.OrdinalIgnoreCase)
                || string.Equals(paragraph.LineSpacingRule, "Exact", StringComparison.OrdinalIgnoreCase)
                || string.Equals(paragraph.LineSpacingRule, "AtLeast", StringComparison.OrdinalIgnoreCase)) {
                return Math.Round((paragraph.LineSpacingValue.Value / 240d) * 100d, 2, MidpointRounding.AwayFromZero);
            }

            return null;
        }

        private static GoogleDocsParagraph ConvertParagraph(
            WordParagraphSnapshot paragraph,
            bool startsNewSectionBefore = false,
            string? sectionBreakType = null) {
            var converted = new GoogleDocsParagraph {
                Text = paragraph.Text,
                StyleId = paragraph.StyleId,
                StyleName = paragraph.StyleName,
                StartsNewSectionBefore = startsNewSectionBefore,
                SectionBreakType = sectionBreakType,
                IsListItem = paragraph.IsListItem,
                IsOrderedList = paragraph.IsOrderedList,
                ListLevel = paragraph.ListLevel,
                ListStyleName = paragraph.ListStyleName,
                Alignment = paragraph.Alignment,
                IndentStartPoints = paragraph.IndentStartPoints,
                IndentEndPoints = paragraph.IndentEndPoints,
                IndentFirstLinePoints = paragraph.IndentFirstLinePoints,
                SpaceAbovePoints = paragraph.SpaceAbovePoints,
                SpaceBelowPoints = paragraph.SpaceBelowPoints,
                LineSpacingPercent = ResolveLineSpacingPercent(paragraph),
                ShadingFillColorHex = paragraph.ShadingFillColorHex,
                LeftBorder = ConvertParagraphBorder(paragraph.LeftBorder),
                RightBorder = ConvertParagraphBorder(paragraph.RightBorder),
                TopBorder = ConvertParagraphBorder(paragraph.TopBorder),
                BottomBorder = ConvertParagraphBorder(paragraph.BottomBorder),
                IsRightToLeft = paragraph.IsRightToLeft,
                KeepWithNext = paragraph.KeepWithNext,
                KeepLinesTogether = paragraph.KeepLinesTogether,
                AvoidWidowAndOrphan = paragraph.AvoidWidowAndOrphan,
                PageBreakBefore = paragraph.PageBreakBefore,
                BookmarkName = paragraph.BookmarkName,
                BookmarkId = paragraph.BookmarkId,
            };

            foreach (var run in paragraph.Runs) {
                converted.AddRun(new GoogleDocsParagraphRun {
                    Text = run.Text,
                    Bold = run.Bold,
                    Italic = run.Italic,
                    Underline = run.Underline,
                    Strike = run.Strike,
                    FontSize = run.FontSize,
                    FontFamily = run.FontFamily,
                    ForegroundColorHex = run.ColorHex,
                    HighlightColor = run.HighlightColor,
                    VerticalTextAlignment = run.VerticalTextAlignment,
                    CapsStyle = run.CapsStyle,
                    Link = run.IsHyperlink ? new GoogleDocsLink {
                        Uri = run.HyperlinkUri,
                        Anchor = run.HyperlinkAnchor,
                    } : null,
                    Footnote = run.Footnote == null ? null : ConvertFootnote(run.Footnote),
                    InlineImage = run.InlineImage == null ? null : new GoogleDocsInlineImage {
                        FilePath = run.InlineImage.FilePath,
                        FileName = run.InlineImage.FileName,
                        ContentType = run.InlineImage.ContentType,
                        Bytes = run.InlineImage.Bytes,
                        Description = run.InlineImage.Description,
                        Title = run.InlineImage.Title,
                        Width = run.InlineImage.Width,
                        Height = run.InlineImage.Height,
                        IsInline = run.InlineImage.IsInline,
                        WrapText = run.InlineImage.WrapText,
                    },
                });
            }

            foreach (var tabStop in paragraph.TabStops) {
                converted.AddTabStop(new GoogleDocsTabStop {
                    Alignment = tabStop.Alignment,
                    Leader = tabStop.Leader,
                    OffsetPoints = tabStop.PositionPoints,
                });
            }

            return converted;
        }

        private static GoogleDocsFootnote ConvertFootnote(WordFootnoteSnapshot footnote) {
            var converted = new GoogleDocsFootnote {
                ReferenceId = footnote.ReferenceId,
            };

            foreach (var paragraph in footnote.Paragraphs) {
                converted.AddParagraph(ConvertParagraph(paragraph));
            }

            return converted;
        }

        private static GoogleDocsTableCellBorder? ConvertBorder(WordTableCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            return new GoogleDocsTableCellBorder {
                Style = border.Style,
                ColorHex = border.ColorHex,
                Size = border.Size,
            };
        }

        private static GoogleDocsParagraphBorder? ConvertParagraphBorder(WordParagraphBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            return new GoogleDocsParagraphBorder {
                Style = border.Style,
                ColorHex = border.ColorHex,
                Size = border.Size,
                Space = border.Space,
            };
        }

        private static bool HasSectionLayout(WordSectionSnapshot section) {
            return section.PageWidthPoints.HasValue
                || section.PageHeightPoints.HasValue
                || section.MarginTopPoints.HasValue
                || section.MarginBottomPoints.HasValue
                || section.MarginLeftPoints.HasValue
                || section.MarginRightPoints.HasValue
                || section.HeaderMarginPoints.HasValue
                || section.FooterMarginPoints.HasValue
                || (section.ColumnCount.HasValue && section.ColumnCount.Value > 1)
                || section.ColumnSpacingPoints.HasValue
                || section.HasColumnSeparator
                || section.DifferentFirstPage
                || section.PageNumberStart.HasValue
                || string.Equals(section.Orientation, "Landscape", StringComparison.OrdinalIgnoreCase);
        }

        private static GoogleDocsSectionStyle? ConvertSectionStyle(WordSectionSnapshot section) {
            if (!HasSectionLayout(section)) {
                return null;
            }

            return new GoogleDocsSectionStyle {
                Orientation = section.Orientation,
                PageWidthPoints = section.PageWidthPoints,
                PageHeightPoints = section.PageHeightPoints,
                MarginTopPoints = section.MarginTopPoints,
                MarginBottomPoints = section.MarginBottomPoints,
                MarginLeftPoints = section.MarginLeftPoints,
                MarginRightPoints = section.MarginRightPoints,
                HeaderMarginPoints = section.HeaderMarginPoints,
                FooterMarginPoints = section.FooterMarginPoints,
                ColumnCount = section.ColumnCount,
                ColumnSpacingPoints = section.ColumnSpacingPoints,
                HasColumnSeparator = section.HasColumnSeparator,
                UseFirstPageHeaderFooter = section.DifferentFirstPage,
                PageNumberStart = section.PageNumberStart,
            };
        }

        private static GoogleDocsTable ConvertTable(WordTableSnapshot table) {
            var converted = new GoogleDocsTable {
                RowCount = table.RowCount,
                ColumnCount = table.ColumnCount,
                StyleName = table.StyleName,
                Title = table.Title,
                Description = table.Description,
                RepeatHeaderRow = table.RepeatHeaderRow,
                HasHorizontalMerges = table.HasHorizontalMerges,
                HasVerticalMerges = table.HasVerticalMerges,
            };

            foreach (var width in table.ColumnWidthPoints) {
                converted.AddColumnWidth(width);
            }

            foreach (var row in table.Rows) {
                var convertedRow = new GoogleDocsTableRow {
                    RowIndex = row.RowIndex,
                };

                foreach (var cell in row.Cells) {
                    var convertedCell = new GoogleDocsTableCell {
                        ColumnIndex = cell.ColumnIndex,
                        ColumnSpan = Math.Max(1, cell.ColumnSpan),
                        RowSpan = Math.Max(1, cell.RowSpan),
                        ShadingFillColorHex = cell.ShadingFillColorHex,
                        LeftBorder = ConvertBorder(cell.LeftBorder),
                        RightBorder = ConvertBorder(cell.RightBorder),
                        TopBorder = ConvertBorder(cell.TopBorder),
                        BottomBorder = ConvertBorder(cell.BottomBorder),
                        HasHorizontalMerge = cell.HasHorizontalMerge,
                        HasVerticalMerge = cell.HasVerticalMerge,
                    };

                    foreach (var paragraph in cell.Paragraphs) {
                        convertedCell.AddParagraph(ConvertParagraph(paragraph));
                    }

                    convertedRow.AddCell(convertedCell);
                }

                converted.AddRow(convertedRow);
            }

            return converted;
        }
    }
}
