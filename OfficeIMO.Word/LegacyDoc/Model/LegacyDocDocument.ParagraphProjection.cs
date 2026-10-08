using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word.LegacyDoc.Model {
    public sealed partial class LegacyDocDocument {
        /// <summary>Projects a non-body story through the shared paragraph/table parser.</summary>
        internal static IReadOnlyList<LegacyDocBodyBlock> BuildStoryBlocks(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition) {
            var projection = new LegacyDocDocument();
            projection.BuildFormattedParagraphs(characters, formattingRanges, paragraphFormattingRanges,
                Array.Empty<LegacyDocSection>(), bookmarkProjection, picturesByCharacterPosition,
                reportUnsupportedFeatures: false);
            return projection._bodyBlocks;
        }

        private string BuildFormattedParagraphs(
            IReadOnlyList<LegacyDocTextCharacter> characters,
            IReadOnlyList<LegacyDocCharacterFormatRange> formattingRanges,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphFormattingRanges,
            IReadOnlyList<LegacyDocSection> sections,
            LegacyDocBookmarkProjectionTracker bookmarkProjection,
            IReadOnlyDictionary<int, LegacyDocPicture> picturesByCharacterPosition,
            bool reportUnsupportedFeatures) {
            var bodyText = new System.Text.StringBuilder(characters.Count);
            var currentRuns = new List<LegacyDocTextRun>();
            var runText = new System.Text.StringBuilder();
            var runCharacterPositions = new List<int>();
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            LegacyDocHyperlinkTarget currentHyperlinkTarget = default;
            bool hasCurrentRun = false;
            bool inTable = false;
            bool justClosedCell = false;
            var tableRows = new List<LegacyDocTableRow>();
            var currentTableRow = new List<LegacyDocTableCell>();
            var currentTableCellParagraphs = new List<LegacyDocTableCellParagraph>();
            int nextSectionIndex = sections.Count > 1 ? 1 : sections.Count;
            bool reportedUnprojectedSectionBoundary = false;
            int currentParagraphStartCharacter = characters.Count == 0 ? 0 : characters[0].CharacterPosition;
            int? currentTableStartCharacter = null;
            int? currentTableRowStartCharacter = null;
            IReadOnlyList<LegacyDocBookmark>? currentTableRowBoundaryBookmarks = null;

            for (int characterIndex = 0; characterIndex < characters.Count; characterIndex++) {
                LegacyDocTextCharacter textCharacter = characters[characterIndex];
                if (LegacyDocPictureReader.TryCreatePictureRun(
                    textCharacter,
                    picturesByCharacterPosition,
                    GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset),
                    out LegacyDocTextRun? pictureRun)) {
                    FlushRun();
                    currentRuns.Add(pictureRun!);
                    if (inTable) {
                        justClosedCell = false;
                    }

                    continue;
                }

                if (LegacyDocField.TryReadHyperlink(
                    characters,
                    characterIndex,
                    out LegacyDocHyperlinkTarget hyperlinkTarget,
                    out int resultStartIndex,
                    out int resultEndIndex,
                    out int fieldEndIndex)) {
                    AppendHyperlinkResult(hyperlinkTarget, resultStartIndex, resultEndIndex);
                    characterIndex = fieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadPageNumber(
                    characters,
                    characterIndex,
                    out int pageNumberResultStartIndex,
                    out int pageNumberResultEndIndex,
                    out int pageNumberFieldEndIndex)) {
                    AppendPageNumberResult(pageNumberResultStartIndex, pageNumberResultEndIndex);
                    characterIndex = pageNumberFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadNumberOfPages(
                    characters,
                    characterIndex,
                    out int numberOfPagesResultStartIndex,
                    out int numberOfPagesResultEndIndex,
                    out int numberOfPagesFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.NumPages, fieldInstruction: null, numberOfPagesResultStartIndex, numberOfPagesResultEndIndex);
                    characterIndex = numberOfPagesFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadSectionPages(characters, characterIndex,
                    out string sectionPagesInstruction, out int sectionPagesResultStartIndex,
                    out int sectionPagesResultEndIndex, out int sectionPagesFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.SectionPages, sectionPagesInstruction, sectionPagesResultStartIndex, sectionPagesResultEndIndex);
                    characterIndex = sectionPagesFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDateTimeField(
                    characters,
                    characterIndex,
                    out LegacyDocFieldKind dateTimeFieldKind,
                    out string dateInstruction,
                    out int dateResultStartIndex,
                    out int dateResultEndIndex,
                    out int dateFieldEndIndex)) {
                    AppendFieldResult(dateTimeFieldKind, dateInstruction, dateResultStartIndex, dateResultEndIndex);
                    characterIndex = dateFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDocumentPropertyField(
                    characters,
                    characterIndex,
                    out string documentPropertyInstruction,
                    out int documentPropertyResultStartIndex,
                    out int documentPropertyResultEndIndex,
                    out int documentPropertyFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.DocumentProperty, documentPropertyInstruction, documentPropertyResultStartIndex, documentPropertyResultEndIndex);
                    characterIndex = documentPropertyFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadEquationField(
                    characters,
                    characterIndex,
                    out string equationInstruction,
                    out int equationResultStartIndex,
                    out int equationResultEndIndex,
                    out int equationFieldEndIndex)) {
                    AppendFieldResult(LegacyDocFieldKind.Equation, equationInstruction, equationResultStartIndex, equationResultEndIndex);
                    characterIndex = equationFieldEndIndex;
                    continue;
                }

                if (LegacyDocField.TryReadDisplayField(
                    characters,
                    characterIndex,
                    out int fallbackFieldResultStartIndex,
                    out int fallbackFieldResultEndIndex,
                    out int fallbackFieldEndIndex)) {
                    AppendFieldDisplayResult(fallbackFieldResultStartIndex, fallbackFieldResultEndIndex);
                    characterIndex = fallbackFieldEndIndex;
                    continue;
                }

                if (textCharacter.Character == '\a') {
                    LegacyDocParagraphFormat paragraphFormat = GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset);
                    LegacyDocCharacterFormat paragraphMarkFormat = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                    if (paragraphMarkFormat.HasFormatting) {
                        paragraphFormat = paragraphFormat.WithParagraphMarkFormat(paragraphMarkFormat);
                    }

                    if (paragraphFormat.IsInTable == true) {
                        if (paragraphFormat.IsTableTerminatingParagraph == true) {
                            AddCurrentTableRow(paragraphFormat, textCharacter.CharacterPosition + 1);
                            // The next paragraph's table depth ends a real table.
                            // Its content, including an empty body paragraph, belongs outside it.
                            bool followingTableParagraph = characterIndex + 1 < characters.Count &&
                                GetParagraphFormatForFileOffset(paragraphFormattingRanges, characters[characterIndex + 1].FileOffset).IsInTable == true;
                            if (!followingTableParagraph) {
                                FlushTable(LegacyDocParagraphFormat.Default, textCharacter.CharacterPosition + 1);
                                currentParagraphStartCharacter = textCharacter.CharacterPosition + 1;
                            }
                        } else {
                            AddCurrentTextAsTableCell(paragraphFormat, allowHeuristicRowTerminator: false);
                        }
                    } else {
                        AddCurrentTextAsTableCell(paragraphFormat, allowHeuristicRowTerminator: true);
                    }

                    AddSectionBreaksAtBodyBoundary(textCharacter.CharacterPosition + 1);
                    currentParagraphStartCharacter = textCharacter.CharacterPosition + 1;
                    continue;
                }

                // 0x0C is also a section mark when the next section starts immediately
                // after it. Project that boundary as a paragraph, not an inline page break.
                bool isSectionMark = textCharacter.Character == LegacyDocSpecialCharacters.PageBreak &&
                    nextSectionIndex < sections.Count &&
                    sections[nextSectionIndex].StartCharacter == textCharacter.CharacterPosition + 1;
                char? normalized = isSectionMark ? '\r' : NormalizeBodyCharacter(textCharacter.Character);
                if (normalized == null) {
                    continue;
                }

                LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, textCharacter.FileOffset);
                if (normalized.Value == '\r') {
                    LegacyDocParagraphFormat paragraphFormat = GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset);
                    if (format.HasFormatting) {
                        paragraphFormat = paragraphFormat.WithParagraphMarkFormat(format);
                    }

                    if (paragraphFormat.IsInTable == true && paragraphFormat.IsTableTerminatingParagraph != true) {
                        AddCurrentTextAsTableCellParagraph(paragraphFormat);
                    } else if (inTable) {
                        FlushTable(GetParagraphFormatForFileOffset(paragraphFormattingRanges, textCharacter.FileOffset), textCharacter.CharacterPosition + 1);
                    } else if (!isSectionMark || currentRuns.Count > 0 || runText.Length > 0 ||
                        !paragraphFormat.Equals(LegacyDocParagraphFormat.Default) &&
                        !paragraphFormat.Equals(new LegacyDocParagraphFormat(null, styleIndex: 0)) ||
                        Bookmarks.Any(bookmark =>
                            bookmark.StartCharacter >= currentParagraphStartCharacter && bookmark.StartCharacter <= textCharacter.CharacterPosition + 1 ||
                            bookmark.EndCharacter >= currentParagraphStartCharacter && bookmark.EndCharacter <= textCharacter.CharacterPosition + 1)) {
                        AddCurrentTextAsParagraph(paragraphFormat, isSectionMark);
                    }

                    bodyText.Append('\r');
                    AddSectionBreaksAtBodyBoundary(textCharacter.CharacterPosition + 1);
                    currentParagraphStartCharacter = textCharacter.CharacterPosition + 1;
                    continue;
                }

                AppendRunCharacter(normalized.Value, format, textCharacter.CharacterPosition);
                bodyText.Append(normalized.Value);
            }

            if (inTable) {
                FlushTable(LegacyDocParagraphFormat.Default, characters.Count == 0 ? 0 : characters[characters.Count - 1].CharacterPosition + 1);
            } else if (currentRuns.Count > 0 || runText.Length > 0) {
                AddCurrentTextAsParagraph(LegacyDocParagraphFormat.Default);
            }

            ReportRemainingUnprojectedSectionBoundaries();
            return bodyText.ToString();

            void AddSectionBreaksAtBodyBoundary(int characterPosition) {
                while (nextSectionIndex < sections.Count && sections[nextSectionIndex].StartCharacter < characterPosition) {
                    ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    nextSectionIndex++;
                }

                while (nextSectionIndex < sections.Count && sections[nextSectionIndex].StartCharacter == characterPosition) {
                    if (IsInsideActiveTable()) {
                        ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    } else if (characterPosition < characters.Count) {
                        _bodyBlocks.Add(new LegacyDocSectionBreakBlock(sections[nextSectionIndex].Format));
                    }

                    nextSectionIndex++;
                }
            }

            bool IsInsideActiveTable() =>
                inTable
                || tableRows.Count > 0
                || currentTableRow.Count > 0
                || currentTableCellParagraphs.Count > 0;

            void ReportRemainingUnprojectedSectionBoundaries() {
                while (nextSectionIndex < sections.Count) {
                    if (sections[nextSectionIndex].StartCharacter < characters.Count) {
                        ReportUnprojectedSectionBoundary(sections[nextSectionIndex].StartCharacter);
                    }

                    nextSectionIndex++;
                }
            }

            void ReportUnprojectedSectionBoundary(int characterPosition) {
                if (reportedUnprojectedSectionBoundary) {
                    return;
                }

                reportedUnprojectedSectionBoundary = true;
                AddUnsupportedFeature(new LegacyDocUnsupportedFeature(
                    LegacyDocUnsupportedFeatureKind.Section,
                    "DOC-MULTIPLE-SECTIONS-PRESENT",
                    $"The legacy DOC contains a section boundary at character position {characterPosition} that does not align with a supported body-block boundary. That section is preserved in the source file but is not projected into the OfficeIMO document.",
                    detailCode: "Fib:PlcfSed"),
                    reportUnsupportedFeatures);
            }

            void AppendHyperlinkResult(LegacyDocHyperlinkTarget hyperlinkTarget, int resultStartIndex, int resultEndIndex) {
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    if (LegacyDocField.TryReadEquationField(
                        characters,
                        resultIndex,
                        out string equationInstruction,
                        out int equationResultStartIndex,
                        out int equationResultEndIndex,
                        out int equationFieldEndIndex) &&
                        equationFieldEndIndex < resultEndIndex) {
                        FlushRun();
                        LegacyDocCharacterFormat equationFormat = LegacyDocCharacterFormat.Default;
                        var equationPositions = new List<int>();
                        var equationText = new System.Text.StringBuilder();
                        foreach (int equationResultIndex in LegacyDocField.EnumerateVisibleResultIndexes(characters, equationResultStartIndex, equationResultEndIndex)) {
                            LegacyDocTextCharacter equationCharacter = characters[equationResultIndex];
                            char? normalizedEquationCharacter = NormalizeBodyCharacter(equationCharacter.Character);
                            if (normalizedEquationCharacter == null) continue;
                            if (equationPositions.Count == 0) {
                                equationFormat = GetFormatForFileOffset(formattingRanges, equationCharacter.FileOffset);
                            }
                            equationText.Append(normalizedEquationCharacter.Value);
                            equationPositions.Add(equationCharacter.CharacterPosition);
                            bodyText.Append(normalizedEquationCharacter.Value);
                        }
                        currentRuns.Add(LegacyDocTextRunFactory.CreateFieldRun(
                            equationText.ToString(),
                            LegacyDocFieldKind.Equation,
                            equationInstruction,
                            equationFormat,
                            equationPositions,
                            hyperlinkTarget));
                        resultIndex = equationFieldEndIndex;
                        continue;
                    }

                    if (characters[resultIndex].Character == LegacyDocField.Begin &&
                        LegacyDocField.TryReadNestedFieldResult(
                            characters,
                            resultIndex,
                            resultEndIndex,
                            out int nestedResultStartIndex,
                            out int nestedResultEndIndex,
                            out int nestedFieldEndIndex)) {
                        foreach (int visibleResultIndex in LegacyDocField.EnumerateVisibleResultIndexes(
                            characters,
                            nestedResultStartIndex,
                            nestedResultEndIndex)) {
                            LegacyDocTextCharacter visibleCharacter = characters[visibleResultIndex];
                            char? normalizedVisibleCharacter = NormalizeBodyCharacter(visibleCharacter.Character);
                            if (normalizedVisibleCharacter == null) continue;
                            LegacyDocCharacterFormat visibleFormat = GetFormatForFileOffset(formattingRanges, visibleCharacter.FileOffset);
                            AppendRunCharacter(normalizedVisibleCharacter.Value, visibleFormat, visibleCharacter.CharacterPosition, hyperlinkTarget);
                            bodyText.Append(normalizedVisibleCharacter.Value);
                        }
                        resultIndex = nestedFieldEndIndex;
                        continue;
                    }

                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    AppendRunCharacter(normalized.Value, format, resultCharacter.CharacterPosition, hyperlinkTarget);
                    bodyText.Append(normalized.Value);
                }
            }

            void AppendPageNumberResult(int resultStartIndex, int resultEndIndex) {
                AppendFieldResult(LegacyDocFieldKind.Page, fieldInstruction: null, resultStartIndex, resultEndIndex);
            }

            void AppendFieldDisplayResult(int resultStartIndex, int resultEndIndex) {
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    LegacyDocCharacterFormat format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    AppendRunCharacter(normalized.Value, format, resultCharacter.CharacterPosition);
                    bodyText.Append(normalized.Value);
                }
            }

            void AppendFieldResult(LegacyDocFieldKind fieldKind, string? fieldInstruction, int resultStartIndex, int resultEndIndex) {
                FlushRun();
                LegacyDocCharacterFormat format = LegacyDocCharacterFormat.Default;
                var positions = new List<int>();
                var resultText = new System.Text.StringBuilder();
                for (int resultIndex = resultStartIndex; resultIndex < resultEndIndex; resultIndex++) {
                    LegacyDocTextCharacter resultCharacter = characters[resultIndex];
                    char? normalized = NormalizeBodyCharacter(resultCharacter.Character);
                    if (normalized == null) {
                        continue;
                    }

                    if (positions.Count == 0) {
                        format = GetFormatForFileOffset(formattingRanges, resultCharacter.FileOffset);
                    }

                    resultText.Append(normalized.Value);
                    positions.Add(resultCharacter.CharacterPosition);
                }

                currentRuns.Add(LegacyDocTextRunFactory.CreateFieldRun(
                    fieldKind == LegacyDocFieldKind.Page ? string.Empty : resultText.ToString(),
                    fieldKind,
                    fieldInstruction,
                    format,
                    positions));
                if (inTable) {
                    justClosedCell = false;
                }
            }

            void AppendRunCharacter(char character, LegacyDocCharacterFormat format, int characterPosition, LegacyDocHyperlinkTarget hyperlinkTarget = default) {
                if (!hasCurrentRun
                    || !format.Equals(currentFormat)
                    || hyperlinkTarget != currentHyperlinkTarget) {
                    FlushRun();
                    currentFormat = format;
                    currentHyperlinkTarget = hyperlinkTarget;
                    hasCurrentRun = true;
                }

                runText.Append(character);
                runCharacterPositions.Add(characterPosition);
                if (inTable) {
                    justClosedCell = false;
                }
            }

            void FlushRun() {
                if (runText.Length == 0) {
                    return;
                }

                currentRuns.Add(new LegacyDocTextRun(
                    runText.ToString(),
                    currentFormat.Bold,
                    currentFormat.Italic,
                    currentFormat.Strike,
                    currentFormat.DoubleStrike,
                    currentFormat.Outline,
                    currentFormat.Shadow,
                    currentFormat.Emboss,
                    currentFormat.Imprint,
                    currentFormat.Hidden,
                    currentFormat.NoProof,
                    currentFormat.Caps,
                    currentFormat.VerticalPosition,
                    currentFormat.Underline,
                    currentFormat.Highlight,
                    currentFormat.FontSizeHalfPoints,
                    currentFormat.ColorHex,
                    currentFormat.FontFamily,
                    runCharacterPositions,
                    currentHyperlinkTarget.Uri,
                    currentHyperlinkTarget.Anchor,
                    specified: currentFormat.Specified,
                    styleRelative: currentFormat.StyleRelative,
                    styleInverted: currentFormat.StyleInverted,
                    characterSpacingTwips: currentFormat.CharacterSpacingTwips,
                    characterScalePercentage: currentFormat.CharacterScalePercentage,
                    kerningMinimumFontSizeHalfPoints: currentFormat.KerningMinimumFontSizeHalfPoints,
                    language: currentFormat.Language,
                    eastAsiaLanguage: currentFormat.EastAsiaLanguage,
                    revision: currentFormat.Revision,
                    hyperlinkTooltip: currentHyperlinkTarget.Tooltip,
                    hyperlinkTargetFrame: currentHyperlinkTarget.TargetFrame));
                runText.Clear();
                runCharacterPositions.Clear();
                currentHyperlinkTarget = default;
            }

            void AddCurrentTextAsParagraph(LegacyDocParagraphFormat paragraphFormat, bool endsWithSectionMark = false) {
                FlushRun();
                IReadOnlyList<LegacyDocTextRun> runs = currentRuns.ToArray();
                _paragraphTextRuns.Add(runs);
                _paragraphFormats.Add(paragraphFormat);
                _paragraphs.Add(string.Concat(runs.Select(run => run.Text)));
                int paragraphEndCharacter = Math.Max(currentParagraphStartCharacter + runs.Sum(run => run.Text.Length), GetRunEndCharacter(runs));
                _bodyBlocks.Add(new LegacyDocParagraphBlock(
                    runs,
                    paragraphFormat,
                    currentParagraphStartCharacter,
                    paragraphEndCharacter,
                    bookmarkProjection.ExtractProjectedParagraphBookmarks(currentParagraphStartCharacter, paragraphEndCharacter),
                    endsWithSectionMark));
                currentRuns.Clear();
                hasCurrentRun = false;
            }

            void AddCurrentTextAsTableCellParagraph(LegacyDocParagraphFormat paragraphFormat) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = false;
            }

            void AddCurrentTextAsTableCell(LegacyDocParagraphFormat paragraphFormat, bool allowHeuristicRowTerminator) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    justClosedCell = false;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                if (allowHeuristicRowTerminator && currentRuns.Count == 0 && justClosedCell) {
                    if (currentTableRow.Count > 0) {
                        tableRows.Add(new LegacyDocTableRow(currentTableRow.ToArray(), bookmarksBefore: currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks()));
                        currentTableRow.Clear();
                    }

                    currentTableRowBoundaryBookmarks = null;
                    justClosedCell = false;
                    return;
                }

                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                    currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                }

                currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                currentTableCellParagraphs.Clear();
                currentRuns.Clear();
                hasCurrentRun = false;
                justClosedCell = true;
            }

            void AddCurrentTableRow(LegacyDocParagraphFormat paragraphFormat, int rowEndCharacter) {
                FlushRun();
                if (!inTable) {
                    inTable = true;
                    currentTableStartCharacter = currentParagraphStartCharacter;
                    currentTableRowStartCharacter = currentParagraphStartCharacter;
                }

                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count > 0 || (!justClosedCell && currentTableRow.Count == 0)) {
                    if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                        currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                    }

                    currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                    currentTableCellParagraphs.Clear();
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(
                        currentTableRow.ToArray(),
                        paragraphFormat.TableCellWidthsTwips,
                        paragraphFormat.TableLeftIndentTwips,
                        paragraphFormat.TableRowHeightTwips,
                        paragraphFormat.TableRowHeightIsExact,
                        paragraphFormat.TableRowCantSplit,
                        paragraphFormat.TableRowIsHeader,
                        paragraphFormat.TableAlignment,
                        paragraphFormat.TableCellHorizontalMerges,
                        paragraphFormat.TableCellVerticalMerges,
                        paragraphFormat.TableCellVerticalAlignments,
                        paragraphFormat.TableCellTextDirections,
                        paragraphFormat.TableCellFitTexts,
                        paragraphFormat.TableCellNoWraps,
                        paragraphFormat.TableCellHideMarks,
                        paragraphFormat.GetTableCellMarginsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellShadingsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellBordersForCellCount(currentTableRow.Count),
                        paragraphFormat.DefaultTableCellSpacingTwips,
                        paragraphFormat.TablePreferredWidth,
                        paragraphFormat.TableAutofit,
                        currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks(),
                        paragraphFormat.TableStyleIndex,
                        paragraphFormat.TableBorders));
                    currentTableRow.Clear();
                }

                hasCurrentRun = false;
                justClosedCell = false;
                currentTableRowBoundaryBookmarks = null;
                currentTableRowStartCharacter = rowEndCharacter;
            }

            void FlushTable(LegacyDocParagraphFormat paragraphFormat, int tableEndCharacter) {
                FlushRun();
                if (currentRuns.Count > 0 || currentTableCellParagraphs.Count > 0) {
                    if (currentRuns.Count > 0 || currentTableCellParagraphs.Count == 0) {
                        currentTableCellParagraphs.Add(CreateCurrentTableCellParagraph(paragraphFormat));
                    }

                    currentTableRow.Add(new LegacyDocTableCell(currentTableCellParagraphs.ToArray()));
                    currentTableCellParagraphs.Clear();
                    currentRuns.Clear();
                }

                if (currentTableRow.Count > 0) {
                    tableRows.Add(new LegacyDocTableRow(
                        currentTableRow.ToArray(),
                        paragraphFormat.TableCellWidthsTwips,
                        paragraphFormat.TableLeftIndentTwips,
                        paragraphFormat.TableRowHeightTwips,
                        paragraphFormat.TableRowHeightIsExact,
                        paragraphFormat.TableRowCantSplit,
                        paragraphFormat.TableRowIsHeader,
                        paragraphFormat.TableAlignment,
                        paragraphFormat.TableCellHorizontalMerges,
                        paragraphFormat.TableCellVerticalMerges,
                        paragraphFormat.TableCellVerticalAlignments,
                        paragraphFormat.TableCellTextDirections,
                        paragraphFormat.TableCellFitTexts,
                        paragraphFormat.TableCellNoWraps,
                        paragraphFormat.TableCellHideMarks,
                        paragraphFormat.GetTableCellMarginsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellShadingsForCellCount(currentTableRow.Count),
                        paragraphFormat.GetTableCellBordersForCellCount(currentTableRow.Count),
                        paragraphFormat.DefaultTableCellSpacingTwips,
                        paragraphFormat.TablePreferredWidth,
                        paragraphFormat.TableAutofit,
                        currentTableRowBoundaryBookmarks ?? ExtractCurrentTableRowBoundaryBookmarks(),
                        paragraphFormat.TableStyleIndex,
                        paragraphFormat.TableBorders));
                    currentTableRow.Clear();
                }

                if (tableRows.Count > 0) {
                    int tableStartCharacter = currentTableStartCharacter ?? GetTableStartCharacter(tableRows);
                    _bodyBlocks.Add(new LegacyDocTableBlock(
                        tableRows.ToArray(),
                        tableStartCharacter,
                        tableEndCharacter,
                        bookmarkProjection.ExtractUnprojectedBlockBookmarks(tableStartCharacter, tableEndCharacter)));
                    tableRows.Clear();
                }

                hasCurrentRun = false;
                inTable = false;
                justClosedCell = false;
                currentTableStartCharacter = null;
                currentTableRowStartCharacter = null;
                currentTableRowBoundaryBookmarks = null;
            }

            LegacyDocTableCellParagraph CreateCurrentTableCellParagraph(LegacyDocParagraphFormat paragraphFormat) {
                IReadOnlyList<LegacyDocTextRun> runs = currentRuns.ToArray();
                int paragraphStartCharacter = GetRunStartCharacter(runs);
                int paragraphEndCharacter = GetRunEndCharacter(runs);
                if (currentTableRowBoundaryBookmarks == null
                    && currentTableRowStartCharacter.HasValue
                    && paragraphStartCharacter == currentTableRowStartCharacter.Value) {
                    currentTableRowBoundaryBookmarks = ExtractCurrentTableRowBoundaryBookmarks();
                }

                return new LegacyDocTableCellParagraph(
                    runs,
                    paragraphFormat,
                    paragraphStartCharacter,
                    paragraphEndCharacter,
                    bookmarkProjection.ExtractProjectedParagraphBookmarks(paragraphStartCharacter, paragraphEndCharacter));
            }

            IReadOnlyList<LegacyDocBookmark> ExtractCurrentTableRowBoundaryBookmarks() {
                int rowStartCharacter = currentTableRowStartCharacter ?? currentTableStartCharacter ?? currentParagraphStartCharacter;
                return rowStartCharacter != currentTableStartCharacter
                    ? bookmarkProjection.ExtractZeroLengthBoundaryBookmarks(rowStartCharacter)
                    : Array.Empty<LegacyDocBookmark>();
            }

            int GetTableStartCharacter(IReadOnlyList<LegacyDocTableRow> rows) {
                foreach (LegacyDocTableRow row in rows) {
                    foreach (LegacyDocTableCell cell in row.Cells) {
                        foreach (LegacyDocTableCellParagraph paragraph in cell.Paragraphs) {
                            return paragraph.StartCharacter;
                        }
                    }
                }

                return currentParagraphStartCharacter;
            }

            int GetRunStartCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
                foreach (LegacyDocTextRun run in runs) {
                    if (run.CharacterPositions.Count > 0) {
                        return run.CharacterPositions[0];
                    }
                }

                return currentParagraphStartCharacter;
            }

            int GetRunEndCharacter(IReadOnlyList<LegacyDocTextRun> runs) {
                for (int index = runs.Count - 1; index >= 0; index--) {
                    IReadOnlyList<int> positions = runs[index].CharacterPositions;
                    if (positions.Count > 0) {
                        return positions[positions.Count - 1] + 1;
                    }
                }

                return currentParagraphStartCharacter;
            }
        }

        private static char? NormalizeBodyCharacter(char character) {
            switch (character) {
                case '\0':
                case '\a':
                    return null;
                case LegacyDocFootnoteReader.FootnoteReferenceCharacter:
                case LegacyDocCommentReader.CommentReferenceCharacter:
                case LegacyDocSpecialCharacters.TextWrappingBreak:
                case LegacyDocSpecialCharacters.PageBreak:
                case LegacyDocSpecialCharacters.ColumnBreak:
                    return character;
                case '\n':
                    return '\r';
                default:
                    if (!char.IsControl(character) || character == '\t' || character == '\r' || character == '\n') {
                        return character;
                    }

                    return null;
            }
        }

        private static LegacyDocCharacterFormat GetFormatForFileOffset(IReadOnlyList<LegacyDocCharacterFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocCharacterFormat.Default;
        }

        private static LegacyDocParagraphFormat GetParagraphFormatForFileOffset(IReadOnlyList<LegacyDocParagraphFormatRange> ranges, int fileOffset) {
            for (int i = 0; i < ranges.Count; i++) {
                if (ranges[i].Contains(fileOffset)) {
                    return ranges[i].Format;
                }
            }

            return LegacyDocParagraphFormat.Default;
        }

    }
}
