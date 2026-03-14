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
                        "Word section header/footer metadata is preserved in the snapshot. Default header/footer execution is implemented, while first/even variants still remain a dedicated follow-up slice.");
                }

                AddSegment(batch, report, section.Index, section.DefaultHeader);
                AddSegment(batch, report, section.Index, section.DefaultFooter);
                AddUnsupportedSegmentNotice(report, section.Index, section.FirstHeader);
                AddUnsupportedSegmentNotice(report, section.Index, section.FirstFooter);
                AddUnsupportedSegmentNotice(report, section.Index, section.EvenHeader);
                AddUnsupportedSegmentNotice(report, section.Index, section.EvenFooter);

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

                        batch.Add(new GoogleDocsInsertParagraphRequest {
                            SectionIndex = sectionIndex,
                            ElementIndex = elementIndex,
                            Paragraph = ConvertParagraph(paragraph, startsNewSectionBefore, section.SectionBreakType),
                        });
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
                            Table = ConvertTable(table),
                            StartsNewSectionBefore = startsNewSectionBefore,
                            SectionBreakType = section.SectionBreakType,
                        });
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

        private static void AddUnsupportedSegmentNotice(
            TranslationReport report,
            int sectionIndex,
            WordHeaderFooterSnapshot? source) {
            if (source == null) {
                return;
            }

            report.Add(
                TranslationSeverity.Warning,
                "HeadersAndFooters",
                $"Section {sectionIndex + 1} contains a {source.Kind} {source.Variant} variant, but the current Google Docs slice only executes default headers and footers.");
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
                || paragraph.PageBreakBefore
                || paragraph.Runs.Any(run =>
                    run.Bold
                    || run.Italic
                    || run.Underline
                    || run.Strike
                    || run.FontSize.HasValue
                    || !string.IsNullOrWhiteSpace(run.ColorHex));
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
                PageBreakBefore = paragraph.PageBreakBefore,
            };

            foreach (var run in paragraph.Runs) {
                converted.AddRun(new GoogleDocsParagraphRun {
                    Text = run.Text,
                    Bold = run.Bold,
                    Italic = run.Italic,
                    Underline = run.Underline,
                    Strike = run.Strike,
                    FontSize = run.FontSize,
                    ForegroundColorHex = run.ColorHex,
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
