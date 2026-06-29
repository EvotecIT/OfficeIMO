using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Loads a legacy binary `.doc` document and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        public static WordDocument LoadLegacyDoc(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return ProjectLoadedLegacyDocDocument(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` document and returns both the projected OfficeIMO document and import report.
        /// </summary>
        public static LegacyDocLoadResult LoadLegacyDocWithReport(string path, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(path, options);
            return CreateLegacyDocLoadResult(document, path);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and projects supported content into a normal OfficeIMO Word document.
        /// The resulting document saves through the normal Open XML path.
        /// </summary>
        public static WordDocument LoadLegacyDoc(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return ProjectLoadedLegacyDocDocument(document, sourcePath: null);
        }

        /// <summary>
        /// Loads a legacy binary `.doc` stream and returns both the projected OfficeIMO document and import report.
        /// </summary>
        public static LegacyDocLoadResult LoadLegacyDocWithReport(Stream stream, LegacyDocImportOptions? options = null) {
            LegacyDocDocument document = LegacyDocDocument.Load(stream, options);
            return CreateLegacyDocLoadResult(document, sourcePath: null);
        }

        private static WordDocument LoadLegacyDocFromNormalFlow(byte[] bytes, string? sourcePath, bool autoSave) {
            if (autoSave) {
                throw new NotSupportedException("Auto-save is not supported when loading legacy binary .doc files. Load the document, then save explicitly to a .docx path.");
            }

            LegacyDocDocument document = LegacyDocDocument.Load(bytes, new LegacyDocImportOptions());
            LegacyDocImportDiagnostic[] errors = document.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            return ProjectLoadedLegacyDocDocument(document, sourcePath);
        }

        private static LegacyDocLoadResult CreateLegacyDocLoadResult(LegacyDocDocument legacyDocument, string? sourcePath) {
            try {
                return new LegacyDocLoadResult(ProjectLoadedLegacyDocDocument(legacyDocument, sourcePath), legacyDocument);
            } catch (InvalidDataException exception) {
                return new LegacyDocLoadResult(document: null, legacyDocument, exception);
            }
        }

        private static WordDocument ProjectLoadedLegacyDocDocument(LegacyDocDocument legacyDocument, string? sourcePath) {
            LegacyDocImportDiagnostic[] errors = legacyDocument.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyDocDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy DOC import failed: " + FormatLegacyDocDiagnostics(errors));
            }

            WordDocument document = CreateInternal(filePath: null, stream: null, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, autoSave: false);
            ApplyLegacyDocProperties(document, legacyDocument.DocumentProperties);
            AddLegacyDocParagraphStyleDefinitions(document, legacyDocument.StyleSheet);
            WordSection section = document.Sections.Count > 0
                ? document.Sections[0]
                : new WordSection(document, null!, null!);
            ApplyLegacyDocSectionFormatting(section, legacyDocument.SectionFormat);

            if (legacyDocument.BodyBlocks.Count == 0) {
                section.AddParagraph();
            } else {
                foreach (LegacyDocBodyBlock block in legacyDocument.BodyBlocks) {
                    if (block is LegacyDocParagraphBlock paragraphBlock) {
                        AddLegacyDocParagraph(section, paragraphBlock.Runs, paragraphBlock.Format, legacyDocument.StyleSheet);
                    } else if (block is LegacyDocTableBlock tableBlock) {
                        AddLegacyDocTable(section, tableBlock, legacyDocument.StyleSheet);
                    }
                }
            }

            document.MarkLoadedFromLegacyDoc(sourcePath, legacyDocument);
            return document;
        }

        private static void ApplyLegacyDocSectionFormatting(WordSection section, LegacyDocSectionFormat sectionFormat) {
            if (!sectionFormat.HasFormatting) {
                return;
            }

            if (sectionFormat.Orientation != null) {
                section.PageOrientation = sectionFormat.Orientation.Value;
            }

            if (sectionFormat.PageWidthTwips != null) {
                section.PageSettings.Width = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.PageWidthTwips.Value;
            }

            if (sectionFormat.PageHeightTwips != null) {
                section.PageSettings.Height = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.PageHeightTwips.Value;
            }

            if (sectionFormat.MarginTopTwips != null) {
                section.Margins.Top = sectionFormat.MarginTopTwips.Value;
            }

            if (sectionFormat.MarginRightTwips != null) {
                section.Margins.Right = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.MarginRightTwips.Value;
            }

            if (sectionFormat.MarginBottomTwips != null) {
                section.Margins.Bottom = sectionFormat.MarginBottomTwips.Value;
            }

            if (sectionFormat.MarginLeftTwips != null) {
                section.Margins.Left = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.MarginLeftTwips.Value;
            }

            PageMargin? pageMargin = section._sectionProperties.GetFirstChild<PageMargin>();
            if (pageMargin == null && (sectionFormat.HeaderDistanceTwips != null || sectionFormat.FooterDistanceTwips != null || sectionFormat.GutterTwips != null)) {
                section._sectionProperties.Append(new PageMargin {
                    Top = 1440,
                    Right = 1440U,
                    Bottom = 1440,
                    Left = 1440U,
                    Header = 720U,
                    Footer = 720U,
                    Gutter = 0U
                });
                pageMargin = section._sectionProperties.GetFirstChild<PageMargin>();
            }

            if (pageMargin != null) {
                if (sectionFormat.HeaderDistanceTwips != null) {
                    pageMargin.Header = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.HeaderDistanceTwips.Value;
                }

                if (sectionFormat.FooterDistanceTwips != null) {
                    pageMargin.Footer = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.FooterDistanceTwips.Value;
                }

                if (sectionFormat.GutterTwips != null) {
                    pageMargin.Gutter = (DocumentFormat.OpenXml.UInt32Value)(uint)sectionFormat.GutterTwips.Value;
                }
            }
        }

        private static void AddLegacyDocTable(WordSection section, LegacyDocTableBlock tableBlock, LegacyDocStyleSheet styleSheet) {
            int rowCount = tableBlock.Rows.Count;
            int columnCount = tableBlock.Rows.Count == 0
                ? 0
                : tableBlock.Rows.Max(row => row.Cells.Count);
            if (rowCount == 0 || columnCount == 0) {
                return;
            }

            WordTable table = section.AddTable(rowCount, columnCount, WordTableStyle.TableGrid);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < sourceRow.Cells.Count && columnIndex < columnCount; columnIndex++) {
                    AddLegacyDocTableCell(table.Rows[rowIndex].Cells[columnIndex], sourceRow.Cells[columnIndex], styleSheet);
                }
            }
        }

        private static void AddLegacyDocTableCell(WordTableCell cell, LegacyDocTableCell sourceCell, LegacyDocStyleSheet styleSheet) {
            if (sourceCell.Runs.Count == 0) {
                WordParagraph emptyParagraph = cell.AddParagraph(string.Empty, removeExistingParagraphs: true);
                ApplyLegacyDocParagraphFormatting(emptyParagraph, sourceCell.Format, styleSheet);
                return;
            }

            LegacyDocTextRun firstRun = sourceCell.Runs[0];
            WordParagraph paragraph;
            if (!ContainsLegacyDocSpecialRunCharacter(firstRun.Text)) {
                paragraph = cell.AddParagraph(firstRun.Text, removeExistingParagraphs: true);
                ApplyLegacyDocRunFormatting(paragraph, firstRun);
            } else {
                paragraph = cell.AddParagraph(string.Empty, removeExistingParagraphs: true);
                AddLegacyDocRunContent(paragraph, firstRun);
            }

            ApplyLegacyDocParagraphFormatting(paragraph, sourceCell.Format, styleSheet);

            for (int index = 1; index < sourceCell.Runs.Count; index++) {
                AddLegacyDocRunContent(paragraph, sourceCell.Runs[index]);
            }
        }

        private static void AddLegacyDocParagraph(WordSection section, IReadOnlyList<LegacyDocTextRun> paragraphRuns, LegacyDocParagraphFormat paragraphFormat, LegacyDocStyleSheet styleSheet) {
            if (paragraphRuns.Count == 0) {
                ApplyLegacyDocParagraphFormatting(section.AddParagraph(), paragraphFormat, styleSheet);
                return;
            }

            WordParagraph paragraph = section.AddParagraph(string.Empty);
            ApplyLegacyDocParagraphFormatting(paragraph, paragraphFormat, styleSheet);
            AddLegacyDocRuns(paragraph, paragraphRuns);
        }

        private static void AddLegacyDocRuns(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> paragraphRuns) {
            foreach (LegacyDocTextRun legacyRun in paragraphRuns) {
                AddLegacyDocRunContent(paragraph, legacyRun);
            }
        }

        private static void AddLegacyDocRunContent(WordParagraph paragraph, LegacyDocTextRun legacyRun) {
            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (!IsLegacyDocSpecialRunCharacter(character)) {
                    continue;
                }

                AddLegacyDocTextSegment(paragraph, legacyRun, segmentStart, index - segmentStart);
                if (character == '\t') {
                    WordParagraph tabRun = paragraph.AddTab();
                    ApplyLegacyDocRunFormatting(tabRun, legacyRun);
                } else {
                    AddLegacyDocBreak(paragraph, legacyRun, character == '\f' ? BreakValues.Page : null);
                }

                segmentStart = index + 1;
            }

            AddLegacyDocTextSegment(paragraph, legacyRun, segmentStart, text.Length - segmentStart);
        }

        private static bool ContainsLegacyDocSpecialRunCharacter(string text) {
            for (int index = 0; index < text.Length; index++) {
                if (IsLegacyDocSpecialRunCharacter(text[index])) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsLegacyDocSpecialRunCharacter(char character) {
            return character == '\t' || character == '\v' || character == '\f';
        }

        private static void AddLegacyDocBreak(WordParagraph paragraph, LegacyDocTextRun legacyRun, BreakValues? breakType) {
            var run = new Run();
            run.Append(breakType == null ? new Break() : new Break { Type = breakType });
            paragraph._paragraph.Append(run);
            var breakRun = new WordParagraph(paragraph._document, paragraph._paragraph, run);
            ApplyLegacyDocRunFormatting(breakRun, legacyRun);
        }

        private static void AddLegacyDocTextSegment(WordParagraph paragraph, LegacyDocTextRun legacyRun, int startIndex, int length) {
            if (length <= 0) {
                return;
            }

            WordParagraph run = paragraph.AddText(legacyRun.Text.Substring(startIndex, length));
            ApplyLegacyDocRunFormatting(run, legacyRun);
        }

        private static void ApplyLegacyDocParagraphFormatting(WordParagraph paragraph, LegacyDocParagraphFormat paragraphFormat, LegacyDocStyleSheet styleSheet) {
            if (paragraphFormat.StyleIndex != null) {
                ApplyLegacyDocParagraphStyle(paragraph, paragraphFormat.StyleIndex.Value, styleSheet);
            }

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                paragraph.ParagraphAlignment = alignment;
            }

            if (paragraphFormat.SpacingBeforeTwips != null) {
                paragraph.LineSpacingBefore = paragraphFormat.SpacingBeforeTwips;
            }

            if (paragraphFormat.SpacingAfterTwips != null) {
                paragraph.LineSpacingAfter = paragraphFormat.SpacingAfterTwips;
            }

            if (paragraphFormat.LineSpacingTwips != null) {
                paragraph.LineSpacing = paragraphFormat.LineSpacingTwips;
            }

            if (paragraphFormat.LeftIndentTwips != null) {
                paragraph.IndentationBefore = paragraphFormat.LeftIndentTwips;
            }

            if (paragraphFormat.RightIndentTwips != null) {
                paragraph.IndentationAfter = paragraphFormat.RightIndentTwips;
            }

            if (paragraphFormat.FirstLineIndentTwips != null) {
                if (paragraphFormat.FirstLineIndentTwips.Value < 0) {
                    paragraph.IndentationHanging = -paragraphFormat.FirstLineIndentTwips.Value;
                } else {
                    paragraph.IndentationFirstLine = paragraphFormat.FirstLineIndentTwips;
                }
            }

            if (paragraphFormat.KeepLinesTogether == true) {
                paragraph.KeepLinesTogether = true;
            }

            if (paragraphFormat.KeepWithNext == true) {
                paragraph.KeepWithNext = true;
            }

            if (paragraphFormat.PageBreakBefore == true) {
                paragraph.PageBreakBefore = true;
            }

            if (paragraphFormat.AvoidWidowAndOrphan == true) {
                paragraph.AvoidWidowAndOrphan = true;
            }
        }

        private static void ApplyLegacyDocParagraphStyle(WordParagraph paragraph, ushort styleIndex, LegacyDocStyleSheet styleSheet) {
            if (styleSheet.TryGetParagraphStyle(styleIndex, out LegacyDocParagraphStyle legacyStyle)) {
                if (legacyStyle.BuiltInStyle != null) {
                    paragraph.SetStyle(legacyStyle.BuiltInStyle.Value);
                    return;
                }

                if (!string.IsNullOrWhiteSpace(legacyStyle.StyleId)) {
                    paragraph.SetStyleId(legacyStyle.StyleId!);
                    return;
                }
            }

            if (TryMapBuiltInParagraphStyle(styleIndex, out WordParagraphStyles style)) {
                paragraph.SetStyle(style);
            }
        }

        private static void AddLegacyDocParagraphStyleDefinitions(WordDocument document, LegacyDocStyleSheet styleSheet) {
            StyleDefinitionsPart? styleDefinitionsPart = document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart;
            if (styleDefinitionsPart == null) {
                return;
            }

            Styles styles = styleDefinitionsPart.Styles ??= new Styles();
            var existingStyleIds = new HashSet<string>(
                styles.OfType<Style>()
                    .Select(style => style.StyleId?.Value)
                    .Where(styleId => !string.IsNullOrWhiteSpace(styleId))!,
                StringComparer.OrdinalIgnoreCase);

            foreach (LegacyDocParagraphStyle legacyStyle in styleSheet.ParagraphStyles) {
                if (legacyStyle.BuiltInStyle != null || string.IsNullOrWhiteSpace(legacyStyle.StyleId)) {
                    continue;
                }

                if (!existingStyleIds.Add(legacyStyle.StyleId!)) {
                    continue;
                }

                var style = new Style { Type = StyleValues.Paragraph, StyleId = legacyStyle.StyleId, CustomStyle = true };
                style.Append(new StyleName { Val = legacyStyle.Name });
                style.Append(new BasedOn { Val = ResolveLegacyDocBasedOnStyleId(legacyStyle, styleSheet) });
                StyleParagraphProperties? styleParagraphProperties = CreateLegacyDocStyleParagraphProperties(legacyStyle.ParagraphFormat);
                if (styleParagraphProperties != null) {
                    style.Append(styleParagraphProperties);
                }

                StyleRunProperties? styleRunProperties = CreateLegacyDocStyleRunProperties(legacyStyle.CharacterFormat);
                if (styleRunProperties != null) {
                    style.Append(styleRunProperties);
                }

                styles.Append(style);
            }

            styles.Save();
        }

        private static string ResolveLegacyDocBasedOnStyleId(LegacyDocParagraphStyle legacyStyle, LegacyDocStyleSheet styleSheet) {
            if (legacyStyle.BasedOnStyleIndex != null
                && legacyStyle.BasedOnStyleIndex.Value != legacyStyle.Index
                && styleSheet.TryGetParagraphStyle(legacyStyle.BasedOnStyleIndex.Value, out LegacyDocParagraphStyle basedOnStyle)) {
                if (basedOnStyle.BuiltInStyle != null) {
                    return basedOnStyle.BuiltInStyle.Value.ToStringStyle();
                }

                if (!string.IsNullOrWhiteSpace(basedOnStyle.StyleId)) {
                    return basedOnStyle.StyleId!;
                }
            }

            if (legacyStyle.BasedOnStyleIndex != null
                && TryMapBuiltInParagraphStyle(legacyStyle.BasedOnStyleIndex.Value, out WordParagraphStyles builtInStyle)) {
                return builtInStyle.ToStringStyle();
            }

            return WordParagraphStyles.Normal.ToStringStyle();
        }

        private static StyleParagraphProperties? CreateLegacyDocStyleParagraphProperties(LegacyDocParagraphFormat paragraphFormat) {
            var properties = new StyleParagraphProperties();
            bool hasProperties = false;

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                properties.Append(new Justification { Val = alignment });
                hasProperties = true;
            }

            SpacingBetweenLines? spacing = null;
            if (paragraphFormat.SpacingBeforeTwips != null) {
                spacing ??= new SpacingBetweenLines();
                spacing.Before = paragraphFormat.SpacingBeforeTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            if (paragraphFormat.SpacingAfterTwips != null) {
                spacing ??= new SpacingBetweenLines();
                spacing.After = paragraphFormat.SpacingAfterTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            if (paragraphFormat.LineSpacingTwips != null) {
                spacing ??= new SpacingBetweenLines();
                spacing.Line = paragraphFormat.LineSpacingTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                spacing.LineRule = LineSpacingRuleValues.AtLeast;
            }

            if (spacing != null) {
                properties.Append(spacing);
                hasProperties = true;
            }

            Indentation? indentation = null;
            if (paragraphFormat.LeftIndentTwips != null) {
                indentation ??= new Indentation();
                indentation.Left = paragraphFormat.LeftIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            if (paragraphFormat.RightIndentTwips != null) {
                indentation ??= new Indentation();
                indentation.Right = paragraphFormat.RightIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            if (paragraphFormat.FirstLineIndentTwips != null) {
                indentation ??= new Indentation();
                if (paragraphFormat.FirstLineIndentTwips.Value < 0) {
                    indentation.Hanging = (-paragraphFormat.FirstLineIndentTwips.Value).ToString(System.Globalization.CultureInfo.InvariantCulture);
                } else {
                    indentation.FirstLine = paragraphFormat.FirstLineIndentTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            if (indentation != null) {
                properties.Append(indentation);
                hasProperties = true;
            }

            if (paragraphFormat.KeepLinesTogether == true) {
                properties.Append(new KeepLines());
                hasProperties = true;
            }

            if (paragraphFormat.KeepWithNext == true) {
                properties.Append(new KeepNext());
                hasProperties = true;
            }

            if (paragraphFormat.PageBreakBefore == true) {
                properties.Append(new PageBreakBefore());
                hasProperties = true;
            }

            if (paragraphFormat.AvoidWidowAndOrphan == true) {
                properties.Append(new WidowControl());
                hasProperties = true;
            }

            return hasProperties ? properties : null;
        }

        private static StyleRunProperties? CreateLegacyDocStyleRunProperties(LegacyDocCharacterFormat characterFormat) {
            var properties = new StyleRunProperties();
            bool hasProperties = false;

            if (!string.IsNullOrEmpty(characterFormat.FontFamily)) {
                properties.Append(new RunFonts {
                    Ascii = characterFormat.FontFamily,
                    HighAnsi = characterFormat.FontFamily,
                    ComplexScript = characterFormat.FontFamily,
                    EastAsia = characterFormat.FontFamily
                });
                hasProperties = true;
            }

            if (characterFormat.Bold) {
                properties.Append(new Bold());
                hasProperties = true;
            }

            if (characterFormat.Italic) {
                properties.Append(new Italic());
                hasProperties = true;
            }

            if (characterFormat.Strike) {
                properties.Append(new Strike());
                hasProperties = true;
            }

            if (!string.IsNullOrEmpty(characterFormat.ColorHex)) {
                properties.Append(new Color { Val = characterFormat.ColorHex! });
                hasProperties = true;
            }

            if (characterFormat.FontSizeHalfPoints != null) {
                string fontSize = characterFormat.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                properties.Append(new FontSize { Val = fontSize });
                properties.Append(new FontSizeComplexScript { Val = fontSize });
                hasProperties = true;
            }

            if (characterFormat.Highlight != null && TryMapHighlight(characterFormat.Highlight.Value, out HighlightColorValues highlight)) {
                properties.Append(new Highlight { Val = highlight });
                hasProperties = true;
            }

            if (characterFormat.Underline != null && TryMapUnderline(characterFormat.Underline.Value, out UnderlineValues underline)) {
                properties.Append(new Underline { Val = underline });
                hasProperties = true;
            }

            if (characterFormat.VerticalPosition != null && TryMapVerticalPosition(characterFormat.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                properties.Append(new VerticalTextAlignment { Val = verticalPosition });
                hasProperties = true;
            }

            return hasProperties ? properties : null;
        }

        private static bool TryMapBuiltInParagraphStyle(ushort styleIndex, out WordParagraphStyles style) {
            switch (styleIndex) {
                case 0:
                    style = WordParagraphStyles.Normal;
                    return true;
                case 1:
                    style = WordParagraphStyles.Heading1;
                    return true;
                case 2:
                    style = WordParagraphStyles.Heading2;
                    return true;
                case 3:
                    style = WordParagraphStyles.Heading3;
                    return true;
                case 4:
                    style = WordParagraphStyles.Heading4;
                    return true;
                case 5:
                    style = WordParagraphStyles.Heading5;
                    return true;
                case 6:
                    style = WordParagraphStyles.Heading6;
                    return true;
                case 7:
                    style = WordParagraphStyles.Heading7;
                    return true;
                case 8:
                    style = WordParagraphStyles.Heading8;
                    return true;
                case 9:
                    style = WordParagraphStyles.Heading9;
                    return true;
                default:
                    style = default;
                    return false;
            }
        }

        private static void ApplyLegacyDocRunFormatting(WordParagraph run, LegacyDocTextRun legacyRun) {
            if (legacyRun.Bold) {
                run.SetBold();
            }

            if (legacyRun.Italic) {
                run.SetItalic();
            }

            if (legacyRun.Strike) {
                run.SetStrike();
            }

            if (legacyRun.VerticalPosition != null && TryMapVerticalPosition(legacyRun.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                run.VerticalTextAlignment = verticalPosition;
            }

            if (legacyRun.Underline != null && TryMapUnderline(legacyRun.Underline.Value, out UnderlineValues underline)) {
                run.Underline = underline;
            }

            if (legacyRun.Highlight != null && TryMapHighlight(legacyRun.Highlight.Value, out HighlightColorValues highlight)) {
                run.Highlight = highlight;
            }

            if (legacyRun.FontSizeHalfPoints != null) {
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.FontSize = new FontSize {
                    Val = legacyRun.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
                };
            }

            if (!string.IsNullOrEmpty(legacyRun.ColorHex)) {
                run.ColorHex = legacyRun.ColorHex!;
            }

            if (!string.IsNullOrEmpty(legacyRun.FontFamily)) {
                run.SetFontFamily(legacyRun.FontFamily!);
            }
        }

        private static bool TryMapHighlight(LegacyDocHighlightColorKind highlightKind, out HighlightColorValues value) {
            switch (highlightKind) {
                case LegacyDocHighlightColorKind.Black:
                    value = HighlightColorValues.Black;
                    return true;
                case LegacyDocHighlightColorKind.Blue:
                    value = HighlightColorValues.Blue;
                    return true;
                case LegacyDocHighlightColorKind.Cyan:
                    value = HighlightColorValues.Cyan;
                    return true;
                case LegacyDocHighlightColorKind.Green:
                    value = HighlightColorValues.Green;
                    return true;
                case LegacyDocHighlightColorKind.Magenta:
                    value = HighlightColorValues.Magenta;
                    return true;
                case LegacyDocHighlightColorKind.Red:
                    value = HighlightColorValues.Red;
                    return true;
                case LegacyDocHighlightColorKind.Yellow:
                    value = HighlightColorValues.Yellow;
                    return true;
                case LegacyDocHighlightColorKind.White:
                    value = HighlightColorValues.White;
                    return true;
                case LegacyDocHighlightColorKind.DarkBlue:
                    value = HighlightColorValues.DarkBlue;
                    return true;
                case LegacyDocHighlightColorKind.DarkCyan:
                    value = HighlightColorValues.DarkCyan;
                    return true;
                case LegacyDocHighlightColorKind.DarkGreen:
                    value = HighlightColorValues.DarkGreen;
                    return true;
                case LegacyDocHighlightColorKind.DarkMagenta:
                    value = HighlightColorValues.DarkMagenta;
                    return true;
                case LegacyDocHighlightColorKind.DarkRed:
                    value = HighlightColorValues.DarkRed;
                    return true;
                case LegacyDocHighlightColorKind.DarkYellow:
                    value = HighlightColorValues.DarkYellow;
                    return true;
                case LegacyDocHighlightColorKind.DarkGray:
                    value = HighlightColorValues.DarkGray;
                    return true;
                case LegacyDocHighlightColorKind.LightGray:
                    value = HighlightColorValues.LightGray;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static bool TryMapUnderline(LegacyDocUnderlineKind underline, out UnderlineValues value) {
            switch (underline) {
                case LegacyDocUnderlineKind.Single:
                    value = UnderlineValues.Single;
                    return true;
                case LegacyDocUnderlineKind.Words:
                    value = UnderlineValues.Words;
                    return true;
                case LegacyDocUnderlineKind.Double:
                    value = UnderlineValues.Double;
                    return true;
                case LegacyDocUnderlineKind.Dotted:
                    value = UnderlineValues.Dotted;
                    return true;
                case LegacyDocUnderlineKind.Thick:
                    value = UnderlineValues.Thick;
                    return true;
                case LegacyDocUnderlineKind.Dash:
                    value = UnderlineValues.Dash;
                    return true;
                case LegacyDocUnderlineKind.DotDash:
                    value = UnderlineValues.DotDash;
                    return true;
                case LegacyDocUnderlineKind.DotDotDash:
                    value = UnderlineValues.DotDotDash;
                    return true;
                case LegacyDocUnderlineKind.Wave:
                    value = UnderlineValues.Wave;
                    return true;
                case LegacyDocUnderlineKind.DottedHeavy:
                    value = UnderlineValues.DottedHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashedHeavy:
                    value = UnderlineValues.DashedHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashDotHeavy:
                    value = UnderlineValues.DashDotHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashDotDotHeavy:
                    value = UnderlineValues.DashDotDotHeavy;
                    return true;
                case LegacyDocUnderlineKind.WavyHeavy:
                    value = UnderlineValues.WavyHeavy;
                    return true;
                case LegacyDocUnderlineKind.DashLong:
                    value = UnderlineValues.DashLong;
                    return true;
                case LegacyDocUnderlineKind.WavyDouble:
                    value = UnderlineValues.WavyDouble;
                    return true;
                case LegacyDocUnderlineKind.DashLongHeavy:
                    value = UnderlineValues.DashLongHeavy;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static bool TryMapVerticalPosition(LegacyDocVerticalPositionKind position, out VerticalPositionValues value) {
            switch (position) {
                case LegacyDocVerticalPositionKind.Superscript:
                    value = VerticalPositionValues.Superscript;
                    return true;
                case LegacyDocVerticalPositionKind.Subscript:
                    value = VerticalPositionValues.Subscript;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static bool TryMapParagraphAlignment(LegacyDocParagraphAlignment alignment, out JustificationValues value) {
            switch (alignment) {
                case LegacyDocParagraphAlignment.Left:
                    value = JustificationValues.Left;
                    return true;
                case LegacyDocParagraphAlignment.Center:
                    value = JustificationValues.Center;
                    return true;
                case LegacyDocParagraphAlignment.Right:
                    value = JustificationValues.Right;
                    return true;
                case LegacyDocParagraphAlignment.Justify:
                    value = JustificationValues.Both;
                    return true;
                default:
                    value = default;
                    return false;
            }
        }

        private static void ApplyLegacyDocProperties(WordDocument document, LegacyDocDocumentProperties properties) {
            if (!properties.HasAnyProperties) {
                return;
            }

            document.BuiltinDocumentProperties.Title = properties.Title;
            document.BuiltinDocumentProperties.Subject = properties.Subject;
            document.BuiltinDocumentProperties.Creator = properties.Creator;
            document.BuiltinDocumentProperties.Keywords = properties.Keywords;
            document.BuiltinDocumentProperties.Description = properties.Description;
            document.BuiltinDocumentProperties.Category = properties.Category;
            document.BuiltinDocumentProperties.LastModifiedBy = properties.LastModifiedBy;
            document.BuiltinDocumentProperties.Revision = properties.Revision;
            document.BuiltinDocumentProperties.Created = properties.Created;
            document.BuiltinDocumentProperties.Modified = properties.Modified;
            document.BuiltinDocumentProperties.LastPrinted = properties.LastPrinted;

            if (properties.Company != null) {
                document.ApplicationProperties.Company = properties.Company;
            }

            if (properties.Manager != null) {
                document.ApplicationProperties.Manager = new Manager { Text = properties.Manager };
            }

            foreach (KeyValuePair<string, LegacyDocDocumentPropertyValue> property in properties.CustomProperties) {
                if (TryCreateWordCustomProperty(property.Value, out WordCustomProperty? wordProperty)) {
                    document.CustomDocumentProperties[property.Key] = wordProperty!;
                }
            }
        }

        private static bool TryCreateWordCustomProperty(LegacyDocDocumentPropertyValue property, out WordCustomProperty? wordProperty) {
            switch (property.Kind) {
                case LegacyDocDocumentPropertyValueKind.Text:
                    wordProperty = new WordCustomProperty(Convert.ToString(property.Value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
                case LegacyDocDocumentPropertyValueKind.Boolean:
                    wordProperty = new WordCustomProperty(Convert.ToBoolean(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.DateTime:
                    wordProperty = new WordCustomProperty(Convert.ToDateTime(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Integer:
                    wordProperty = new WordCustomProperty(Convert.ToInt32(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                case LegacyDocDocumentPropertyValueKind.Number:
                    wordProperty = new WordCustomProperty(Convert.ToDouble(property.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return true;
                default:
                    wordProperty = null;
                    return false;
            }
        }

        private static string FormatLegacyDocDiagnostics(IEnumerable<LegacyDocImportDiagnostic> diagnostics) {
            const int maxDiagnostics = 6;
            LegacyDocImportDiagnostic[] selected = diagnostics.Take(maxDiagnostics + 1).ToArray();
            string message = string.Join("; ", selected.Take(maxDiagnostics).Select(diagnostic => diagnostic.ToString()));
            if (selected.Length > maxDiagnostics) {
                message += $"; and {selected.Length - maxDiagnostics} more";
            }

            return message;
        }
    }
}
