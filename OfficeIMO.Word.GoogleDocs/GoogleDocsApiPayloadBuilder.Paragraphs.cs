using OfficeIMO.GoogleWorkspace;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.GoogleDocs {
    internal static partial class GoogleDocsApiPayloadBuilder {
        private static void AppendParagraphRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsParagraph paragraph,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId) {
            AppendParagraphRequests(payload, report, paragraph, imageUris, segmentId, 1, true, segmentId == null, null, allowNamedRanges: segmentId == null, sectionStyle: null);
        }

        private static void AppendParagraphRequests(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsParagraph paragraph,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId,
            int insertionIndex,
            bool allowStructuralBreaks,
            bool allowFootnotes,
            List<GoogleDocsFootnote>? preparedFootnotes,
            bool allowNamedRanges,
            GoogleDocsSectionStyle? sectionStyle) {
            var materialized = MaterializeParagraph(paragraph, report, imageUris);
            payload.Requests.Add(new GoogleDocsApiRequestPayload {
                InsertText = new GoogleDocsApiInsertTextRequestPayload {
                    Location = new GoogleDocsApiLocationPayload {
                        Index = insertionIndex,
                        SegmentId = segmentId,
                    },
                    Text = materialized.InsertedText,
                }
            });

            if (!string.IsNullOrWhiteSpace(paragraph.BookmarkName) && allowNamedRanges && materialized.InsertedText.Length > 0) {
                int namedRangeEndIndex = insertionIndex + Math.Max(1, materialized.InsertedText.Length - 1);
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    CreateNamedRange = new GoogleDocsApiCreateNamedRangeRequestPayload {
                        Name = paragraph.BookmarkName!,
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = namedRangeEndIndex,
                            SegmentId = segmentId,
                        }
                    }
                });
            }

            var paragraphFields = new List<string>();
            var paragraphStyle = new GoogleDocsApiParagraphStylePayload();
            if (TryMapNamedStyle(paragraph, out var namedStyleType)) {
                paragraphStyle.NamedStyleType = namedStyleType;
                paragraphFields.Add("namedStyleType");
            }

            if (TryMapAlignment(paragraph.Alignment, out var alignment)) {
                paragraphStyle.Alignment = alignment;
                paragraphFields.Add("alignment");
            }

            if (paragraph.IsRightToLeft) {
                paragraphStyle.Direction = "RIGHT_TO_LEFT";
                paragraphFields.Add("direction");
            }

            if (paragraph.KeepWithNext) {
                paragraphStyle.KeepWithNext = true;
                paragraphFields.Add("keepWithNext");
            }

            if (paragraph.KeepLinesTogether) {
                paragraphStyle.KeepLinesTogether = true;
                paragraphFields.Add("keepLinesTogether");
            }

            if (paragraph.AvoidWidowAndOrphan) {
                paragraphStyle.AvoidWidowAndOrphan = true;
                paragraphFields.Add("avoidWidowAndOrphan");
            }

            if (BuildParagraphTabStops(paragraph.TabStops) is { Count: > 0 } tabStops) {
                paragraphStyle.TabStops = tabStops;
                paragraphFields.Add("tabStops");
            }

            if (TryBuildParagraphDimension(paragraph.IndentStartPoints, out var indentStart)) {
                paragraphStyle.IndentStart = indentStart;
                paragraphFields.Add("indentStart");
            }

            if (TryBuildParagraphDimension(paragraph.IndentEndPoints, out var indentEnd)) {
                paragraphStyle.IndentEnd = indentEnd;
                paragraphFields.Add("indentEnd");
            }

            if (TryBuildParagraphDimension(paragraph.IndentFirstLinePoints, out var indentFirstLine)) {
                paragraphStyle.IndentFirstLine = indentFirstLine;
                paragraphFields.Add("indentFirstLine");
            }

            if (TryBuildParagraphDimension(paragraph.SpaceAbovePoints, out var spaceAbove)) {
                paragraphStyle.SpaceAbove = spaceAbove;
                paragraphFields.Add("spaceAbove");
            }

            if (TryBuildParagraphDimension(paragraph.SpaceBelowPoints, out var spaceBelow)) {
                paragraphStyle.SpaceBelow = spaceBelow;
                paragraphFields.Add("spaceBelow");
            }

            if (paragraph.LineSpacingPercent.HasValue) {
                paragraphStyle.LineSpacing = paragraph.LineSpacingPercent.Value;
                paragraphFields.Add("lineSpacing");
            }

            if (!string.IsNullOrWhiteSpace(paragraph.ShadingFillColorHex)) {
                var shadingColor = BuildOptionalColor(paragraph.ShadingFillColorHex);
                if (shadingColor != null) {
                    paragraphStyle.Shading = new GoogleDocsApiParagraphShadingPayload {
                        BackgroundColor = shadingColor,
                    };
                    paragraphFields.Add("shading");
                }
            }

            if (BuildParagraphBorder(paragraph.LeftBorder) is { } leftBorder) {
                paragraphStyle.BorderLeft = leftBorder;
                paragraphFields.Add("borderLeft");
            }

            if (BuildParagraphBorder(paragraph.RightBorder) is { } rightBorder) {
                paragraphStyle.BorderRight = rightBorder;
                paragraphFields.Add("borderRight");
            }

            if (BuildParagraphBorder(paragraph.TopBorder) is { } topBorder) {
                paragraphStyle.BorderTop = topBorder;
                paragraphFields.Add("borderTop");
            }

            if (BuildParagraphBorder(paragraph.BottomBorder) is { } bottomBorder) {
                paragraphStyle.BorderBottom = bottomBorder;
                paragraphFields.Add("borderBottom");
            }

            if (paragraphFields.Count > 0) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    UpdateParagraphStyle = new GoogleDocsApiUpdateParagraphStyleRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = insertionIndex + materialized.InsertedText.Length,
                            SegmentId = segmentId,
                        },
                        ParagraphStyle = paragraphStyle,
                        Fields = string.Join(",", paragraphFields),
                    }
                });
            }

            AppendSectionStyleRequest(
                payload,
                insertionIndex,
                insertionIndex + materialized.InsertedText.Length,
                segmentId,
                sectionStyle);

            foreach (var run in materialized.Runs) {
                var textStyle = BuildTextStyle(run.Source);
                if (textStyle == null) {
                    continue;
                }

                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    UpdateTextStyle = new GoogleDocsApiUpdateTextStyleRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex + run.StartOffset,
                            EndIndex = insertionIndex + run.EndOffset,
                            SegmentId = segmentId,
                        },
                        TextStyle = textStyle,
                        Fields = BuildTextStyleFields(textStyle),
                    }
                });
            }

            if (paragraph.IsListItem && materialized.InsertedText.Length > 1) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    CreateParagraphBullets = new GoogleDocsApiCreateParagraphBulletsRequestPayload {
                        Range = new GoogleDocsApiRangePayload {
                            StartIndex = insertionIndex,
                            EndIndex = insertionIndex + materialized.InsertedText.Length,
                            SegmentId = segmentId,
                        },
                        BulletPreset = ResolveListPreset(paragraph, report),
                    }
                });
            }

            if (materialized.Footnotes.Count > 0 && allowFootnotes) {
                foreach (var footnote in materialized.Footnotes.OrderByDescending(item => item.InsertOffset)) {
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        CreateFootnote = new GoogleDocsApiCreateFootnoteRequestPayload {
                            Location = new GoogleDocsApiLocationPayload {
                                Index = insertionIndex + footnote.InsertOffset,
                                SegmentId = segmentId,
                            }
                        }
                    });
                    preparedFootnotes?.Add(footnote.Source);
                }
            }

            if (materialized.Footnotes.Count > 0 && !allowFootnotes) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "Footnotes",
                    segmentId == null
                        ? "Footnotes inside table cell replay are not executed in the current Google Docs slice, because native footnote creation is currently limited to top-level body paragraphs."
                        : "Nested footnotes inside Google Docs segments are not executed in the current slice.");
            }

            foreach (var image in materialized.Images.OrderByDescending(item => item.InsertOffset)) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertInlineImage = new GoogleDocsApiInsertInlineImageRequestPayload {
                        Uri = image.Uri,
                        ObjectSize = BuildImageSize(image.Source),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex + image.InsertOffset,
                            SegmentId = segmentId,
                        }
                    }
                });
            }

            if (paragraph.PageBreakBefore && segmentId == null && allowStructuralBreaks) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertPageBreak = new GoogleDocsApiInsertPageBreakRequestPayload {
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex,
                        }
                    }
                });

                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                        "PageBreaks",
                        "Word paragraphs marked with PageBreakBefore are now emitted as native Google Docs insertPageBreak requests in the body flow.");
            }

            if (paragraph.PageBreakBefore && segmentId != null) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    "PageBreakBefore in a Word header/footer paragraph is not executed in the current Google Docs slice because insertPageBreak is only valid in the document body.");
            }

            if (paragraph.PageBreakBefore && !allowStructuralBreaks) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "PageBreakBefore inside a table cell is ignored in the current Google Docs slice because insertPageBreak is only emitted for top-level document flow.");
            }

            if (paragraph.StartsNewSectionBefore && segmentId == null && allowStructuralBreaks) {
                payload.Requests.Add(new GoogleDocsApiRequestPayload {
                    InsertSectionBreak = new GoogleDocsApiInsertSectionBreakRequestPayload {
                        SectionType = ResolveSectionBreakType(paragraph.SectionBreakType, report),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = insertionIndex,
                        }
                    }
                });
            }

            if (paragraph.StartsNewSectionBefore && segmentId != null) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    "HeadersAndFooters",
                    "Section-break markers are ignored inside header/footer segment content because insertSectionBreak is only valid in the document body.");
            }

            if (paragraph.StartsNewSectionBefore && !allowStructuralBreaks) {
                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Warning,
                    segmentId == null ? "Tables" : "HeadersAndFooters",
                    "Section-break markers inside a table cell are ignored in the current Google Docs slice because insertSectionBreak is only emitted for top-level document flow.");
            }
        }

        private static void AppendRequest(
            GoogleDocsApiBatchUpdatePayload payload,
            TranslationReport report,
            GoogleDocsRequest request,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            string? segmentId,
            bool allowFootnotes,
            List<GoogleDocsFootnote>? preparedFootnotes,
            bool allowNamedRanges) {
            switch (request) {
                case GoogleDocsInsertParagraphRequest paragraphRequest:
                    AppendParagraphRequests(payload, report, paragraphRequest.Paragraph, imageUris, segmentId, 1, true, allowFootnotes, preparedFootnotes, allowNamedRanges, paragraphRequest.SectionStyle);
                    break;
                case GoogleDocsInsertTableRequest tableRequest:
                    payload.Requests.Add(new GoogleDocsApiRequestPayload {
                        InsertTable = new GoogleDocsApiInsertTableRequestPayload {
                            Rows = Math.Max(1, tableRequest.Table.RowCount),
                            Columns = Math.Max(1, tableRequest.Table.ColumnCount),
                            Location = new GoogleDocsApiLocationPayload {
                                Index = 1,
                                SegmentId = segmentId,
                            }
                        }
                    });

                    if (tableRequest.StartsNewSectionBefore && segmentId == null) {
                        payload.Requests.Add(new GoogleDocsApiRequestPayload {
                            InsertSectionBreak = new GoogleDocsApiInsertSectionBreakRequestPayload {
                        SectionType = ResolveSectionBreakType(tableRequest.SectionBreakType, report),
                        Location = new GoogleDocsApiLocationPayload {
                            Index = 1,
                        }
                            }
                        });
                    }

                    if (tableRequest.StartsNewSectionBefore && segmentId != null) {
                        AddReportNoticeOnce(
                            report,
                            TranslationSeverity.Warning,
                            "HeadersAndFooters",
                            "Section-break markers are ignored inside header/footer segment content because insertSectionBreak is only valid in the document body.");
                    }

                    AppendSectionStyleRequest(payload, 1, 2, segmentId, tableRequest.SectionStyle);
                    break;
            }
        }


        private static MaterializedParagraph MaterializeParagraph(
            GoogleDocsParagraph paragraph,
            TranslationReport report,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris) {
            var materialized = new MaterializedParagraph();
            materialized.PrefixText = BuildParagraphPrefix(paragraph);
            var currentOffset = materialized.PrefixText.Length;
            foreach (var run in paragraph.Runs) {
                var runText = BuildRunText(run, report, imageUris, materialized, currentOffset);
                if (string.IsNullOrEmpty(runText)) {
                    continue;
                }

                materialized.TextBuilder.Append(runText);
                materialized.Runs.Add(new MaterializedRun {
                    StartOffset = currentOffset,
                    EndOffset = currentOffset + runText.Length,
                    Source = run,
                });
                currentOffset += runText.Length;
            }

            if (materialized.TextBuilder.Length == 0 && !string.IsNullOrWhiteSpace(paragraph.Text)) {
                var paragraphText = SanitizeText(paragraph.Text);
                materialized.TextBuilder.Append(paragraphText);
                materialized.Runs.Add(new MaterializedRun {
                    StartOffset = currentOffset,
                    EndOffset = currentOffset + paragraphText.Length,
                    Source = new GoogleDocsParagraphRun {
                        Text = paragraphText,
                    }
                });
            }

            materialized.TextBuilder.Append('\n');
            return materialized;
        }

        private static GoogleDocsApiTextStylePayload? BuildTextStyle(GoogleDocsParagraphRun source) {
            var style = new GoogleDocsApiTextStylePayload();
            var hasStyle = false;

            if (source.Bold) {
                style.Bold = true;
                hasStyle = true;
            }

            if (source.Italic) {
                style.Italic = true;
                hasStyle = true;
            }

            if (source.Underline) {
                style.Underline = true;
                hasStyle = true;
            }

            if (source.Strike) {
                style.Strikethrough = true;
                hasStyle = true;
            }

            if (source.FontSize.HasValue) {
                style.FontSize = new GoogleDocsApiDimensionPayload {
                    Magnitude = source.FontSize.Value,
                    Unit = "PT",
                };
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.FontFamily)) {
                style.WeightedFontFamily = new GoogleDocsApiWeightedFontFamilyPayload {
                    FontFamily = source.FontFamily!.Trim(),
                };
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.ForegroundColorHex)) {
                style.ForegroundColor = BuildOptionalColor(source.ForegroundColorHex);
                hasStyle = style.ForegroundColor != null;
            }

            if (!string.IsNullOrWhiteSpace(source.HighlightColor)) {
                style.BackgroundColor = BuildHighlightColor(source.HighlightColor);
                hasStyle |= style.BackgroundColor != null;
            }

            if (!string.IsNullOrWhiteSpace(source.VerticalTextAlignment)) {
                style.BaselineOffset = BuildBaselineOffset(source.VerticalTextAlignment);
                hasStyle |= !string.IsNullOrWhiteSpace(style.BaselineOffset);
            }

            if (string.Equals(source.CapsStyle, "SmallCaps", StringComparison.OrdinalIgnoreCase)) {
                style.SmallCaps = true;
                hasStyle = true;
            }

            if (!string.IsNullOrWhiteSpace(source.Link?.Uri)) {
                var linkUri = source.Link!.Uri!;
                style.Link = new GoogleDocsApiLinkPayload {
                    Url = linkUri,
                };
                hasStyle = true;
            }

            return hasStyle ? style : null;
        }

        private static string BuildTextStyleFields(GoogleDocsApiTextStylePayload style) {
            var fields = new List<string>();
            if (style.Bold.HasValue) fields.Add("bold");
            if (style.Italic.HasValue) fields.Add("italic");
            if (style.Underline.HasValue) fields.Add("underline");
            if (style.Strikethrough.HasValue) fields.Add("strikethrough");
            if (style.FontSize != null) fields.Add("fontSize");
            if (style.WeightedFontFamily != null) fields.Add("weightedFontFamily");
            if (style.ForegroundColor != null) fields.Add("foregroundColor");
            if (style.BackgroundColor != null) fields.Add("backgroundColor");
            if (!string.IsNullOrWhiteSpace(style.BaselineOffset)) fields.Add("baselineOffset");
            if (style.SmallCaps.HasValue) fields.Add("smallCaps");
            if (style.Link != null) fields.Add("link");
            return string.Join(",", fields);
        }

        private static string BuildRunText(
            GoogleDocsParagraphRun run,
            TranslationReport report,
            IReadOnlyDictionary<GoogleDocsInlineImage, string>? imageUris,
            MaterializedParagraph materialized,
            int currentOffset) {
            var text = SanitizeText(run.Text);
            if (run.Footnote != null) {
                materialized.Footnotes.Add(new MaterializedFootnote {
                    InsertOffset = currentOffset + text.Length,
                    Source = run.Footnote,
                });
            }

            if (run.InlineImage == null) {
                return text;
            }

            if (imageUris != null && imageUris.TryGetValue(run.InlineImage, out var uploadedUri) && !string.IsNullOrWhiteSpace(uploadedUri)) {
                materialized.Images.Add(new MaterializedImage {
                    InsertOffset = currentOffset + text.Length,
                    Uri = uploadedUri,
                    Source = run.InlineImage,
                });

                AddReportNoticeOnce(
                    report,
                    TranslationSeverity.Info,
                    "InlineImages",
                    "Inline Word images are exported through temporary Drive-backed public URLs so Google Docs can materialize native inline images.");

                return text;
            }

            AddReportNoticeOnce(
                report,
                TranslationSeverity.Info,
                "InlineImages",
                "Inline Word images without an uploaded Drive-backed URI fall back to readable placeholders in Google Docs export.");

            var label = run.InlineImage.Description ?? run.InlineImage.Title;
            if (string.IsNullOrWhiteSpace(label) && !string.IsNullOrWhiteSpace(run.InlineImage.FilePath)) {
                label = Path.GetFileName(run.InlineImage.FilePath);
            }

            var placeholder = string.IsNullOrWhiteSpace(label) ? "[Image]" : "[Image: " + label!.Trim() + "]";
            return string.IsNullOrWhiteSpace(text) ? placeholder : text + placeholder;
        }

        private static string BuildParagraphPrefix(GoogleDocsParagraph paragraph) {
            if (!paragraph.IsListItem) {
                return string.Empty;
            }

            var level = paragraph.ListLevel.GetValueOrDefault();
            if (level <= 0) {
                return string.Empty;
            }

            return new string('\t', level);
        }
    }
}
