using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
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
            var sectionFormats = new List<(WordSection Section, LegacyDocSectionFormat Format)> {
                (section, legacyDocument.SectionFormat)
            };
            LegacyDocNoteProjection notes = LegacyDocNoteProjection.Create(legacyDocument.Footnotes, legacyDocument.Endnotes, legacyDocument.StyleSheet);

            if (legacyDocument.BodyBlocks.Count == 0) {
                section.AddParagraph();
            } else {
                foreach (LegacyDocBodyBlock block in legacyDocument.BodyBlocks) {
                    if (block is LegacyDocParagraphBlock paragraphBlock) {
                        AddLegacyDocParagraph(section, paragraphBlock, legacyDocument.StyleSheet, notes);
                    } else if (block is LegacyDocSectionBreakBlock sectionBreakBlock) {
                        section = document.AddSection(sectionBreakBlock.Format.SectionBreakType ?? SectionMarkValues.NextPage);
                        sectionFormats.Add((section, sectionBreakBlock.Format));
                    } else if (block is LegacyDocTableBlock tableBlock) {
                        AddLegacyDocTable(section, tableBlock, legacyDocument.StyleSheet, notes);
                    }
                }
            }

            foreach ((WordSection targetSection, LegacyDocSectionFormat sectionFormat) in sectionFormats) {
                ApplyLegacyDocSectionFormatting(targetSection, sectionFormat);
            }

            ApplyLegacyDocDocumentOptions(document, legacyDocument);
            AddLegacyDocHeaderFooterStories(document, legacyDocument.HeaderFooterStories, legacyDocument.StyleSheet);
            document.MarkLoadedFromLegacyDoc(sourcePath, legacyDocument);
            return document;
        }

        private static void ApplyLegacyDocDocumentOptions(WordDocument document, LegacyDocDocument legacyDocument) {
            if (!legacyDocument.DifferentOddAndEvenPages) {
                return;
            }

            foreach (WordSection section in document.Sections) {
                section.DifferentOddAndEvenPages = true;
            }
        }

        private static void AddLegacyDocHeaderFooterStories(WordDocument document, IReadOnlyList<LegacyDocHeaderFooterStory> stories, LegacyDocStyleSheet styleSheet) {
            foreach (LegacyDocHeaderFooterStory story in stories) {
                if (story.SectionIndex < 0 || story.SectionIndex >= document.Sections.Count || story.Paragraphs.Count == 0) {
                    continue;
                }

                WordSection section = document.Sections[story.SectionIndex];
                WordHeaderFooter target = story.IsHeader
                    ? section.GetOrCreateHeader(story.Type)
                    : section.GetOrCreateFooter(story.Type);
                foreach (WordParagraph paragraph in target.Paragraphs.ToList()) {
                    paragraph.Remove();
                }

                foreach (LegacyDocHeaderFooterParagraph sourceParagraph in story.Paragraphs) {
                    WordParagraph paragraph = target.AddParagraph(sourceParagraph.Bookmarks.Count == 0 ? sourceParagraph.Text : string.Empty);
                    if (sourceParagraph.Bookmarks.Count > 0) {
                        paragraph._paragraph.RemoveAllChildren<Run>();
                    }

                    ApplyLegacyDocParagraphFormatting(paragraph, sourceParagraph.Format, styleSheet);
                    ReplaceLegacyDocParagraphRuns(
                        paragraph,
                        sourceParagraph.Runs,
                        LegacyDocNoteProjection.Empty,
                        LegacyDocBookmarkProjection.Create(sourceParagraph.Bookmarks, sourceParagraph.StartCharacter, sourceParagraph.EndCharacter));
                }
            }
        }

        private static void ApplyLegacyDocSectionFormatting(WordSection section, LegacyDocSectionFormat sectionFormat) {
            if (!sectionFormat.HasFormatting) {
                return;
            }

            if (sectionFormat.SectionBreakType != null) {
                SectionType? existingSectionType = section._sectionProperties.GetFirstChild<SectionType>();
                existingSectionType?.Remove();
                section._sectionProperties.Append(new SectionType { Val = sectionFormat.SectionBreakType.Value });
            }

            if (sectionFormat.DifferentFirstPage) {
                section.DifferentFirstPage = true;
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

            if (sectionFormat.ColumnCount != null) {
                section.ColumnCount = sectionFormat.ColumnCount.Value;
            }

            if (sectionFormat.ColumnSpacingTwips != null) {
                section.ColumnsSpace = sectionFormat.ColumnSpacingTwips.Value;
            }

            if (sectionFormat.HasColumnSeparator) {
                section.HasColumnSeparator = true;
            }

            if (sectionFormat.RtlGutter) {
                section.RtlGutter = true;
            }

            if (sectionFormat.VerticalAlignment != null) {
                VerticalTextAlignmentOnPage? verticalAlignment = section._sectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>();
                verticalAlignment?.Remove();
                section._sectionProperties.Append(new VerticalTextAlignmentOnPage { Val = sectionFormat.VerticalAlignment.Value });
            }

            if (sectionFormat.PageBorders != null && sectionFormat.PageBorders.Value.HasAny) {
                ApplyLegacyDocSectionPageBorders(section, sectionFormat.PageBorders.Value);
            }

            if (sectionFormat.LineNumberCountBy != null
                || sectionFormat.LineNumberDistanceTwips != null
                || sectionFormat.LineNumberStart != null
                || sectionFormat.LineNumberRestart != null) {
                LineNumberType? lineNumbering = section._sectionProperties.GetFirstChild<LineNumberType>();
                lineNumbering?.Remove();

                var projectedLineNumbering = new LineNumberType();
                if (sectionFormat.LineNumberCountBy != null) {
                    projectedLineNumbering.CountBy = (short)sectionFormat.LineNumberCountBy.Value;
                }

                if (sectionFormat.LineNumberDistanceTwips != null) {
                    projectedLineNumbering.Distance = sectionFormat.LineNumberDistanceTwips.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (sectionFormat.LineNumberStart != null) {
                    projectedLineNumbering.Start = (short)sectionFormat.LineNumberStart.Value;
                }

                if (sectionFormat.LineNumberRestart != null) {
                    projectedLineNumbering.Restart = sectionFormat.LineNumberRestart.Value;
                }

                section._sectionProperties.Append(projectedLineNumbering);
            }

            if (sectionFormat.PageNumberStart != null || sectionFormat.PageNumberFormat != null) {
                section.AddPageNumbering(sectionFormat.PageNumberStart, sectionFormat.PageNumberFormat);
            }

            if (sectionFormat.FootnotePosition != null
                || sectionFormat.FootnoteRestart != null
                || sectionFormat.FootnoteStart != null
                || sectionFormat.FootnoteNumberFormat != null) {
                section.AddFootnoteProperties(
                    sectionFormat.FootnoteNumberFormat,
                    sectionFormat.FootnotePosition,
                    sectionFormat.FootnoteRestart,
                    sectionFormat.FootnoteStart);
            }

            if (sectionFormat.EndnotePosition != null
                || sectionFormat.EndnoteRestart != null
                || sectionFormat.EndnoteStart != null
                || sectionFormat.EndnoteNumberFormat != null) {
                section.AddEndnoteProperties(
                    numberingFormat: sectionFormat.EndnoteNumberFormat,
                    position: sectionFormat.EndnotePosition,
                    restartNumbering: sectionFormat.EndnoteRestart,
                    startNumber: sectionFormat.EndnoteStart);
            }
        }

        private static void ApplyLegacyDocSectionPageBorders(WordSection section, LegacyDocParagraphBorders borders) {
            PageBorders? pageBorders = section._sectionProperties.GetFirstChild<PageBorders>();
            pageBorders?.Remove();

            pageBorders = new PageBorders();
            if (borders.PageOptions.HasNonDefault) {
                pageBorders.Display = GetLegacyDocPageBorderDisplay(borders.PageOptions.Display);
                pageBorders.OffsetFrom = GetLegacyDocPageBorderOffsetFrom(borders.PageOptions.OffsetFrom);
                pageBorders.ZOrder = GetLegacyDocPageBorderZOrder(borders.PageOptions.ZOrder);
            }

            if (borders.Top.HasAny) {
                pageBorders.TopBorder = CreateLegacyDocParagraphBorder<TopBorder>(borders.Top);
            }

            if (borders.Left.HasAny) {
                pageBorders.LeftBorder = CreateLegacyDocParagraphBorder<LeftBorder>(borders.Left);
            }

            if (borders.Bottom.HasAny) {
                pageBorders.BottomBorder = CreateLegacyDocParagraphBorder<BottomBorder>(borders.Bottom);
            }

            if (borders.Right.HasAny) {
                pageBorders.RightBorder = CreateLegacyDocParagraphBorder<RightBorder>(borders.Right);
            }

            if (pageBorders.ChildElements.Count > 0) {
                section._sectionProperties.Append(pageBorders);
            }
        }

        private static PageBorderDisplayValues GetLegacyDocPageBorderDisplay(LegacyDocPageBorderDisplay display) {
            switch (display) {
                case LegacyDocPageBorderDisplay.AllPages:
                    return PageBorderDisplayValues.AllPages;
                case LegacyDocPageBorderDisplay.FirstPage:
                    return PageBorderDisplayValues.FirstPage;
                case LegacyDocPageBorderDisplay.NotFirstPage:
                    return PageBorderDisplayValues.NotFirstPage;
                default:
                    return PageBorderDisplayValues.AllPages;
            }
        }

        private static PageBorderOffsetValues GetLegacyDocPageBorderOffsetFrom(LegacyDocPageBorderOffsetFrom offsetFrom) {
            switch (offsetFrom) {
                case LegacyDocPageBorderOffsetFrom.Text:
                    return PageBorderOffsetValues.Text;
                case LegacyDocPageBorderOffsetFrom.Page:
                    return PageBorderOffsetValues.Page;
                default:
                    return PageBorderOffsetValues.Text;
            }
        }

        private static PageBorderZOrderValues GetLegacyDocPageBorderZOrder(LegacyDocPageBorderZOrder zOrder) {
            switch (zOrder) {
                case LegacyDocPageBorderZOrder.Front:
                    return PageBorderZOrderValues.Front;
                case LegacyDocPageBorderZOrder.Back:
                    return PageBorderZOrderValues.Back;
                default:
                    return PageBorderZOrderValues.Front;
            }
        }

        private static void AddLegacyDocTable(WordSection section, LegacyDocTableBlock tableBlock, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes) {
            int rowCount = tableBlock.Rows.Count;
            int columnCount = tableBlock.Rows.Count == 0
                ? 0
                : tableBlock.Rows.Max(row => row.Cells.Count);
            if (rowCount == 0 || columnCount == 0) {
                return;
            }

            WordTable table = section.AddTable(rowCount, columnCount, WordTableStyle.TableNormal);
            LegacyDocTableAlignment? tableAlignment = tableBlock.Rows
                .Select(row => row.TableAlignment)
                .FirstOrDefault(alignment => alignment.HasValue);
            if (tableAlignment != null) {
                ApplyLegacyDocTableAlignment(table, tableAlignment.Value);
            }

            int? tableLeftIndentTwips = tableBlock.Rows
                .Select(row => row.TableLeftIndentTwips)
                .FirstOrDefault(indent => indent.HasValue);
            if (tableLeftIndentTwips != null) {
                ApplyLegacyDocTableIndentation(table, tableLeftIndentTwips.Value);
            }

            LegacyDocTablePreferredWidth? tablePreferredWidth = tableBlock.Rows
                .Select(row => row.TablePreferredWidth)
                .FirstOrDefault(width => width.HasValue);
            if (tablePreferredWidth != null) {
                ApplyLegacyDocTablePreferredWidth(table, tablePreferredWidth.Value);
            }

            bool? tableAutofit = tableBlock.Rows
                .Select(row => row.TableAutofit)
                .FirstOrDefault(autofit => autofit.HasValue);
            if (tableAutofit != null) {
                table.LayoutType = tableAutofit.Value ? TableLayoutValues.Autofit : TableLayoutValues.Fixed;
            }

            int? tableCellSpacingTwips = tableBlock.Rows
                .Select(row => row.DefaultCellSpacingTwips)
                .FirstOrDefault(spacing => spacing.HasValue);
            if (tableCellSpacingTwips != null) {
                table.StyleDetails!.CellSpacing = checked((short)tableCellSpacingTwips.Value);
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                ApplyLegacyDocTableRowFormatting(table.Rows[rowIndex], sourceRow);
                for (int columnIndex = 0; columnIndex < sourceRow.Cells.Count && columnIndex < columnCount; columnIndex++) {
                    AddLegacyDocTableCell(table.Rows[rowIndex].Cells[columnIndex], sourceRow.Cells[columnIndex], styleSheet, notes);
                    if (columnIndex < sourceRow.CellWidthsTwips.Count) {
                        table.Rows[rowIndex].Cells[columnIndex].WidthType = TableWidthUnitValues.Dxa;
                        table.Rows[rowIndex].Cells[columnIndex].Width = sourceRow.CellWidthsTwips[columnIndex];
                    }

                    if (columnIndex < sourceRow.CellHorizontalMerges.Count) {
                        ApplyLegacyDocTableCellHorizontalMerge(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellHorizontalMerges[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellVerticalMerges.Count) {
                        ApplyLegacyDocTableCellVerticalMerge(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellVerticalMerges[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellVerticalAlignments.Count) {
                        ApplyLegacyDocTableCellVerticalAlignment(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellVerticalAlignments[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellTextDirections.Count) {
                        ApplyLegacyDocTableCellTextDirection(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellTextDirections[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellFitTexts.Count && sourceRow.CellFitTexts[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].FitText = true;
                    }

                    if (columnIndex < sourceRow.CellNoWraps.Count && sourceRow.CellNoWraps[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].WrapText = false;
                    }

                    if (columnIndex < sourceRow.CellHideMarks.Count && sourceRow.CellHideMarks[columnIndex]) {
                        table.Rows[rowIndex].Cells[columnIndex].HideMark = true;
                    }

                    if (columnIndex < sourceRow.CellMargins.Count) {
                        ApplyLegacyDocTableCellMargins(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellMargins[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellShadings.Count) {
                        ApplyLegacyDocTableCellShading(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellShadings[columnIndex]);
                    }

                    if (columnIndex < sourceRow.CellBorders.Count) {
                        ApplyLegacyDocTableCellBorders(table.Rows[rowIndex].Cells[columnIndex], sourceRow.CellBorders[columnIndex]);
                    }
                }
            }

            AddLegacyDocTableRowBoundaryBookmarks(table, tableBlock);
            AddLegacyDocTableBlockBookmarks(table, tableBlock);
        }

        private static void AddLegacyDocTableRowBoundaryBookmarks(WordTable table, LegacyDocTableBlock tableBlock) {
            int rowCount = Math.Min(table.Rows.Count, tableBlock.Rows.Count);
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                LegacyDocTableRow sourceRow = tableBlock.Rows[rowIndex];
                if (sourceRow.BookmarksBefore.Count == 0) {
                    continue;
                }

                TableRow row = table.Rows[rowIndex]._tableRow;
                foreach (LegacyDocBookmark bookmark in sourceRow.BookmarksBefore.OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                    table._table.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, row);
                    table._table.InsertBefore(new BookmarkEnd { Id = bookmark.ProjectionId }, row);
                }
            }
        }

        private static void AddLegacyDocTableBlockBookmarks(WordTable table, LegacyDocTableBlock tableBlock) {
            if (tableBlock.Bookmarks.Count == 0 || table._table.Parent is not OpenXmlCompositeElement parent) {
                return;
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                parent.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, table._table);
                parent.InsertBefore(new BookmarkEnd { Id = bookmark.ProjectionId }, table._table);
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => !bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.StartCharacter)
                .OrderByDescending(bookmark => bookmark.EndCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                parent.InsertBefore(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, table._table);
            }

            OpenXmlElement afterAnchor = table._table;
            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.EndCharacter == tableBlock.EndCharacter && !bookmark.IsZeroLength)
                .OrderByDescending(bookmark => bookmark.StartCharacter)
                .ThenBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = parent.InsertAfter(new BookmarkEnd { Id = bookmark.ProjectionId }, afterAnchor)!;
            }

            foreach (LegacyDocBookmark bookmark in tableBlock.Bookmarks
                .Where(bookmark => bookmark.IsZeroLength && bookmark.StartCharacter == tableBlock.EndCharacter)
                .OrderBy(bookmark => bookmark.Name, StringComparer.Ordinal)) {
                afterAnchor = parent.InsertAfter(new BookmarkStart { Id = bookmark.ProjectionId, Name = bookmark.Name }, afterAnchor)!;
                afterAnchor = parent.InsertAfter(new BookmarkEnd { Id = bookmark.ProjectionId }, afterAnchor)!;
            }
        }

        private static void ApplyLegacyDocTablePreferredWidth(WordTable table, LegacyDocTablePreferredWidth preferredWidth) {
            switch (preferredWidth.Unit) {
                case LegacyDocTablePreferredWidthUnit.Auto:
                    table.WidthType = TableWidthUnitValues.Auto;
                    table.Width = 0;
                    break;
                case LegacyDocTablePreferredWidthUnit.Percent:
                    table.WidthType = TableWidthUnitValues.Pct;
                    table.Width = preferredWidth.Value;
                    break;
                case LegacyDocTablePreferredWidthUnit.Dxa:
                    table.WidthType = TableWidthUnitValues.Dxa;
                    table.Width = preferredWidth.Value;
                    break;
            }
        }

        private static void ApplyLegacyDocTableAlignment(WordTable table, LegacyDocTableAlignment tableAlignment) {
            switch (tableAlignment) {
                case LegacyDocTableAlignment.Left:
                    table.Alignment = TableRowAlignmentValues.Left;
                    break;
                case LegacyDocTableAlignment.Center:
                    table.Alignment = TableRowAlignmentValues.Center;
                    break;
                case LegacyDocTableAlignment.Right:
                    table.Alignment = TableRowAlignmentValues.Right;
                    break;
            }
        }

        private static void ApplyLegacyDocTableIndentation(WordTable table, int leftIndentTwips) {
            table.CheckTableProperties();
            table._tableProperties!.TableIndentation = new TableIndentation {
                Width = leftIndentTwips,
                Type = TableWidthUnitValues.Dxa
            };
        }

        private static void ApplyLegacyDocTableCellMargins(WordTableCell cell, LegacyDocTableCellMargins margins) {
            if (margins.TopTwips != null) {
                cell.MarginTopWidth = checked((short)margins.TopTwips.Value);
            }

            if (margins.RightTwips != null) {
                cell.MarginRightWidth = checked((short)margins.RightTwips.Value);
            }

            if (margins.BottomTwips != null) {
                cell.MarginBottomWidth = checked((short)margins.BottomTwips.Value);
            }

            if (margins.LeftTwips != null) {
                cell.MarginLeftWidth = checked((short)margins.LeftTwips.Value);
            }
        }

        private static void ApplyLegacyDocTableCellShading(WordTableCell cell, LegacyDocTableCellShading shading) {
            if (!string.IsNullOrEmpty(shading.FillColorHex)) {
                cell.ShadingFillColorHex = shading.FillColorHex!;
            }
        }

        private static void ApplyLegacyDocTableCellBorders(WordTableCell cell, LegacyDocTableCellBorders borders) {
            ApplyLegacyDocTableCellBorder(
                borders.Top,
                style => cell.Borders.TopStyle = style,
                color => cell.Borders.TopColorHex = color,
                size => cell.Borders.TopSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => cell.Borders.TopSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocTableCellBorder(
                borders.Left,
                style => cell.Borders.LeftStyle = style,
                color => cell.Borders.LeftColorHex = color,
                size => cell.Borders.LeftSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => cell.Borders.LeftSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocTableCellBorder(
                borders.Bottom,
                style => cell.Borders.BottomStyle = style,
                color => cell.Borders.BottomColorHex = color,
                size => cell.Borders.BottomSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => cell.Borders.BottomSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocTableCellBorder(
                borders.Right,
                style => cell.Borders.RightStyle = style,
                color => cell.Borders.RightColorHex = color,
                size => cell.Borders.RightSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => cell.Borders.RightSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
        }

        private static void ApplyLegacyDocTableCellBorder(
            LegacyDocTableCellBorder border,
            Action<BorderValues> setStyle,
            Action<string> setColor,
            Action<int> setSize,
            Action<int> setSpace) {
            BorderValues? style = MapLegacyDocTableCellBorderStyle(border.Style);
            if (style == null) {
                return;
            }

            setStyle(style.Value);
            if (!string.IsNullOrEmpty(border.ColorHex)) {
                setColor(border.ColorHex!);
            }

            if (border.SizeEighthPoints > 0) {
                setSize(border.SizeEighthPoints);
            }

            if (border.SpacePoints > 0) {
                setSpace(border.SpacePoints);
            }
        }

        private static BorderValues? MapLegacyDocTableCellBorderStyle(LegacyDocTableCellBorderStyle style) {
            switch (style) {
                case LegacyDocTableCellBorderStyle.Single:
                    return BorderValues.Single;
                case LegacyDocTableCellBorderStyle.Double:
                    return BorderValues.Double;
                case LegacyDocTableCellBorderStyle.Dotted:
                    return BorderValues.Dotted;
                case LegacyDocTableCellBorderStyle.Dashed:
                    return BorderValues.Dashed;
                default:
                    return null;
            }
        }

        private static void ApplyLegacyDocTableCellHorizontalMerge(WordTableCell cell, LegacyDocTableCellHorizontalMerge horizontalMerge) {
            switch (horizontalMerge) {
                case LegacyDocTableCellHorizontalMerge.Restart:
                    cell.HorizontalMerge = MergedCellValues.Restart;
                    break;
                case LegacyDocTableCellHorizontalMerge.Continue:
                    cell.HorizontalMerge = MergedCellValues.Continue;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellVerticalMerge(WordTableCell cell, LegacyDocTableCellVerticalMerge verticalMerge) {
            switch (verticalMerge) {
                case LegacyDocTableCellVerticalMerge.Restart:
                    cell.VerticalMerge = MergedCellValues.Restart;
                    break;
                case LegacyDocTableCellVerticalMerge.Continue:
                    cell.VerticalMerge = MergedCellValues.Continue;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellVerticalAlignment(WordTableCell cell, LegacyDocTableCellVerticalAlignment verticalAlignment) {
            switch (verticalAlignment) {
                case LegacyDocTableCellVerticalAlignment.Center:
                    cell.VerticalAlignment = TableVerticalAlignmentValues.Center;
                    break;
                case LegacyDocTableCellVerticalAlignment.Bottom:
                    cell.VerticalAlignment = TableVerticalAlignmentValues.Bottom;
                    break;
            }
        }

        private static void ApplyLegacyDocTableCellTextDirection(WordTableCell cell, LegacyDocTableCellTextDirection textDirection) {
            switch (textDirection) {
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeft:
                    cell.TextDirection = TextDirectionValues.TopToBottomRightToLeft;
                    break;
                case LegacyDocTableCellTextDirection.BottomToTopLeftToRight:
                    cell.TextDirection = TextDirectionValues.BottomToTopLeftToRight;
                    break;
                case LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated:
                    cell.TextDirection = TextDirectionValues.LefttoRightTopToBottomRotated;
                    break;
                case LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated:
                    cell.TextDirection = TextDirectionValues.TopToBottomRightToLeftRotated;
                    break;
            }
        }

        private static void ApplyLegacyDocTableRowFormatting(WordTableRow row, LegacyDocTableRow sourceRow) {
            if (sourceRow.RowHeightTwips != null) {
                row.AddTableRowProperties();
                TableRowProperties rowProperties = row._tableRow.TableRowProperties!;
                TableRowHeight? rowHeight = rowProperties.GetFirstChild<TableRowHeight>();
                if (rowHeight == null) {
                    rowHeight = new TableRowHeight();
                    rowProperties.InsertAt(rowHeight, 0);
                }

                rowHeight.Val = (uint)sourceRow.RowHeightTwips.Value;
                rowHeight.HeightType = sourceRow.RowHeightIsExact
                    ? HeightRuleValues.Exact
                    : HeightRuleValues.AtLeast;
            }

            if (sourceRow.RowCantSplit == true) {
                row.AllowRowToBreakAcrossPages = false;
            }

            if (sourceRow.RowIsHeader == true) {
                row.RepeatHeaderRowAtTheTopOfEachPage = true;
            }
        }

        private static void AddLegacyDocTableCell(WordTableCell cell, LegacyDocTableCell sourceCell, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes) {
            var pendingBookmarks = new List<LegacyDocBookmark>();
            int pendingBookmarkStartCharacter = int.MaxValue;
            int pendingBookmarkEndCharacter = int.MinValue;
            bool emittedParagraph = false;
            for (int index = 0; index < sourceCell.Paragraphs.Count; index++) {
                LegacyDocTableCellParagraph sourceParagraph = sourceCell.Paragraphs[index];
                if (string.IsNullOrEmpty(sourceParagraph.Text)
                    && sourceParagraph.Bookmarks.Count > 0
                    && index + 1 < sourceCell.Paragraphs.Count) {
                    AddPendingTableCellBookmarks(sourceParagraph.Bookmarks, ref pendingBookmarkStartCharacter, ref pendingBookmarkEndCharacter, pendingBookmarks);
                    continue;
                }

                AddLegacyDocTableCellParagraph(
                    cell,
                    sourceParagraph,
                    styleSheet,
                    notes,
                    removeExistingParagraphs: !emittedParagraph,
                    pendingBookmarks,
                    pendingBookmarkStartCharacter,
                    pendingBookmarkEndCharacter);
                emittedParagraph = true;
                pendingBookmarks.Clear();
                pendingBookmarkStartCharacter = int.MaxValue;
                pendingBookmarkEndCharacter = int.MinValue;
            }
        }

        private static void AddPendingTableCellBookmarks(IReadOnlyList<LegacyDocBookmark> bookmarks, ref int startCharacter, ref int endCharacter, List<LegacyDocBookmark> pendingBookmarks) {
            foreach (LegacyDocBookmark bookmark in bookmarks) {
                pendingBookmarks.Add(bookmark);
                startCharacter = Math.Min(startCharacter, bookmark.StartCharacter);
                endCharacter = Math.Max(endCharacter, bookmark.EndCharacter);
            }
        }

        private static void AddLegacyDocTableCellParagraph(
            WordTableCell cell,
            LegacyDocTableCellParagraph sourceParagraph,
            LegacyDocStyleSheet styleSheet,
            LegacyDocNoteProjection notes,
            bool removeExistingParagraphs,
            IReadOnlyList<LegacyDocBookmark>? pendingBookmarks = null,
            int pendingBookmarkStartCharacter = int.MaxValue,
            int pendingBookmarkEndCharacter = int.MinValue) {
            IReadOnlyList<LegacyDocBookmark> paragraphBookmarks = MergeLegacyDocTableCellBookmarks(sourceParagraph.Bookmarks, pendingBookmarks);
            int paragraphStartCharacter = pendingBookmarks != null && pendingBookmarks.Count > 0
                ? Math.Min(sourceParagraph.StartCharacter, pendingBookmarkStartCharacter)
                : sourceParagraph.StartCharacter;
            int paragraphEndCharacter = pendingBookmarks != null && pendingBookmarks.Count > 0
                ? Math.Max(sourceParagraph.EndCharacter, pendingBookmarkEndCharacter)
                : sourceParagraph.EndCharacter;

            if (sourceParagraph.Runs.Count == 0) {
                WordParagraph emptyParagraph = cell.AddParagraph(removeExistingParagraphs: removeExistingParagraphs);
                emptyParagraph._paragraph.RemoveAllChildren<Run>();
                ApplyLegacyDocParagraphFormatting(emptyParagraph, sourceParagraph.Format, styleSheet);
                LegacyDocBookmarkProjection.Create(paragraphBookmarks, paragraphStartCharacter, paragraphEndCharacter).EmitRemaining(emptyParagraph._paragraph);
                return;
            }

            if (paragraphBookmarks.Count > 0) {
                WordParagraph bookmarkedParagraph = cell.AddParagraph(removeExistingParagraphs: removeExistingParagraphs);
                bookmarkedParagraph._paragraph.RemoveAllChildren<Run>();
                ApplyLegacyDocParagraphFormatting(bookmarkedParagraph, sourceParagraph.Format, styleSheet);
                LegacyDocBookmarkProjection bookmarks = LegacyDocBookmarkProjection.Create(paragraphBookmarks, paragraphStartCharacter, paragraphEndCharacter);
                AddLegacyDocRuns(bookmarkedParagraph, sourceParagraph.Runs, notes, bookmarks);
                bookmarks.EmitRemaining(bookmarkedParagraph._paragraph);
                return;
            }

            LegacyDocTextRun firstRun = sourceParagraph.Runs[0];
            WordParagraph paragraph;
            int remainingRunStartIndex;
            if (ContainsLegacyDocSpecialRunCharacter(firstRun.Text) || firstRun.HyperlinkTarget.HasValue) {
                paragraph = cell.AddParagraph(string.Empty, removeExistingParagraphs: removeExistingParagraphs);
                remainingRunStartIndex = 0;
            } else {
                paragraph = cell.AddParagraph(firstRun.Text, removeExistingParagraphs: removeExistingParagraphs);
                ApplyLegacyDocRunFormatting(paragraph, firstRun);
                remainingRunStartIndex = 1;
            }

            ApplyLegacyDocParagraphFormatting(paragraph, sourceParagraph.Format, styleSheet);
            AddLegacyDocRuns(paragraph, sourceParagraph.Runs, remainingRunStartIndex, notes);
        }

        private static IReadOnlyList<LegacyDocBookmark> MergeLegacyDocTableCellBookmarks(IReadOnlyList<LegacyDocBookmark> paragraphBookmarks, IReadOnlyList<LegacyDocBookmark>? pendingBookmarks) {
            if (pendingBookmarks == null || pendingBookmarks.Count == 0) {
                return paragraphBookmarks;
            }

            if (paragraphBookmarks.Count == 0) {
                return pendingBookmarks;
            }

            return pendingBookmarks
                .Concat(paragraphBookmarks)
                .Distinct()
                .ToArray();
        }

        private static void AddLegacyDocParagraph(WordSection section, LegacyDocParagraphBlock paragraphBlock, LegacyDocStyleSheet styleSheet, LegacyDocNoteProjection notes) {
            IReadOnlyList<LegacyDocTextRun> paragraphRuns = paragraphBlock.Runs;
            LegacyDocParagraphFormat paragraphFormat = paragraphBlock.Format;
            if (paragraphRuns.Count == 0) {
                WordParagraph emptyParagraph = section.AddParagraph();
                ApplyLegacyDocParagraphFormatting(emptyParagraph, paragraphFormat, styleSheet);
                LegacyDocBookmarkProjection.Create(paragraphBlock.Bookmarks, paragraphBlock.StartCharacter, paragraphBlock.EndCharacter).EmitRemaining(emptyParagraph._paragraph);
                return;
            }

            WordParagraph paragraph = section.AddParagraph(string.Empty);
            ApplyLegacyDocParagraphFormatting(paragraph, paragraphFormat, styleSheet);
            LegacyDocBookmarkProjection bookmarks = LegacyDocBookmarkProjection.Create(paragraphBlock.Bookmarks, paragraphBlock.StartCharacter, paragraphBlock.EndCharacter);
            AddLegacyDocRuns(paragraph, paragraphRuns, notes, bookmarks);
            bookmarks.EmitRemaining(paragraph._paragraph);
        }

        private static void AddLegacyDocRuns(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> paragraphRuns, LegacyDocNoteProjection notes) {
            AddLegacyDocRuns(paragraph, paragraphRuns, 0, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void AddLegacyDocRuns(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> paragraphRuns, int startIndex, LegacyDocNoteProjection notes) {
            AddLegacyDocRuns(paragraph, paragraphRuns, startIndex, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void AddLegacyDocRuns(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> paragraphRuns, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            AddLegacyDocRuns(paragraph, paragraphRuns, 0, notes, bookmarks);
        }

        private static void AddLegacyDocRuns(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> paragraphRuns, int startIndex, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            for (int index = startIndex; index < paragraphRuns.Count; index++) {
                LegacyDocTextRun legacyRun = paragraphRuns[index];
                if (legacyRun.IsPageNumber) {
                    AddLegacyDocPageNumber(paragraph, legacyRun, bookmarks);
                    continue;
                }

                if (legacyRun.IsNumPages) {
                    AddLegacyDocNumberOfPages(paragraph, legacyRun, bookmarks);
                    continue;
                }

                if (legacyRun.IsStaticDisplayField) {
                    AddLegacyDocStaticDisplayField(paragraph, legacyRun, bookmarks);
                    continue;
                }

                if (legacyRun.HyperlinkTarget.HasValue) {
                    int hyperlinkStartIndex = index;
                    LegacyDocHyperlinkTarget hyperlinkTarget = legacyRun.HyperlinkTarget;
                    while (index + 1 < paragraphRuns.Count
                        && paragraphRuns[index + 1].HyperlinkTarget == hyperlinkTarget) {
                        index++;
                    }

                    AddLegacyDocHyperlinkRunsContent(paragraph, paragraphRuns, hyperlinkStartIndex, index - hyperlinkStartIndex + 1, notes, bookmarks);
                    continue;
                }

                AddLegacyDocRunContent(paragraph, legacyRun, notes, bookmarks);
            }
        }

        private static void AddLegacyDocRunContent(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes) {
            AddLegacyDocRunContent(paragraph, legacyRun, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void AddLegacyDocRunContent(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            if (legacyRun.IsPageNumber) {
                AddLegacyDocPageNumber(paragraph, legacyRun, bookmarks);
                return;
            }

            if (legacyRun.IsNumPages) {
                AddLegacyDocNumberOfPages(paragraph, legacyRun, bookmarks);
                return;
            }

            if (legacyRun.IsStaticDisplayField) {
                AddLegacyDocStaticDisplayField(paragraph, legacyRun, bookmarks);
                return;
            }

            if (legacyRun.HyperlinkTarget.HasValue) {
                AddLegacyDocHyperlinkRunContent(paragraph, legacyRun, notes, bookmarks);
                return;
            }

            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (!IsLegacyDocSpecialRunCharacter(character)) {
                    int? markerPosition = GetLegacyDocRunCharacterPosition(legacyRun, index);
                    if (!bookmarks.HasMarkers(markerPosition)) {
                        continue;
                    }

                    AddLegacyDocTextSegment(paragraph, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(paragraph._paragraph, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AddLegacyDocTextSegment(paragraph, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, index));
                if (character == '\t') {
                    WordParagraph tabRun = paragraph.AddTab();
                    ApplyLegacyDocRunFormatting(tabRun, legacyRun);
                } else if (character == LegacyDocFootnoteReader.FootnoteReferenceCharacter) {
                    AddLegacyDocNoteReference(paragraph, notes, GetLegacyDocRunCharacterPosition(legacyRun, index));
                } else {
                    AddLegacyDocBreak(paragraph, legacyRun, GetLegacyDocBreakType(character));
                }

                segmentStart = index + 1;
            }

            AddLegacyDocTextSegment(paragraph, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static int? GetLegacyDocRunCharacterPosition(LegacyDocTextRun legacyRun, int index) {
            return index >= 0 && index < legacyRun.CharacterPositions.Count
                ? legacyRun.CharacterPositions[index]
                : null;
        }

        private static int? GetLegacyDocRunEndCharacterPosition(LegacyDocTextRun legacyRun) {
            return legacyRun.CharacterPositions.Count == 0
                ? null
                : legacyRun.CharacterPositions[legacyRun.CharacterPositions.Count - 1] + 1;
        }

        private static void AddLegacyDocNoteReference(WordParagraph paragraph, LegacyDocNoteProjection notes, int? characterPosition) {
            if (characterPosition == null) {
                return;
            }

            if (notes.TryGetFootnote(characterPosition.Value, out LegacyDocFootnote? footnote)) {
                AddLegacyDocFootnoteReference(paragraph, footnote!, notes.StyleSheet);
            } else if (notes.TryGetEndnote(characterPosition.Value, out LegacyDocEndnote? endnote)) {
                AddLegacyDocEndnoteReference(paragraph, endnote!, notes.StyleSheet);
            }
        }

        private static void AddLegacyDocFootnoteReference(WordParagraph paragraph, LegacyDocFootnote footnote, LegacyDocStyleSheet styleSheet) {
            if (footnote.ParagraphRuns.Count == 0) {
                return;
            }

            WordParagraph reference = paragraph.AddFootNote(footnote.ParagraphRuns[0].Text);
            List<WordParagraph>? noteParagraphs = reference.FootNote!.Paragraphs;
            if (noteParagraphs == null || noteParagraphs.Count == 0) {
                return;
            }

            ReplaceLegacyDocNoteParagraphRuns(
                noteParagraphs[0],
                footnote.ParagraphRuns[0],
                keepNoteReferenceMark: true,
                LegacyDocNoteProjection.Empty,
                LegacyDocBookmarkProjection.Create(footnote.ParagraphRuns[0].Bookmarks, footnote.ParagraphRuns[0].StartCharacter, footnote.ParagraphRuns[0].EndCharacter));
            ApplyLegacyDocParagraphFormatting(noteParagraphs[0], footnote.ParagraphRuns[0].Format, styleSheet);
            WordParagraph lastParagraph = noteParagraphs[0];
            for (int index = 1; index < footnote.ParagraphRuns.Count; index++) {
                LegacyDocNoteParagraph sourceParagraph = footnote.ParagraphRuns[index];
                lastParagraph = lastParagraph.AddParagraph(sourceParagraph.Bookmarks.Count == 0 ? sourceParagraph.Text : string.Empty);
                ReplaceLegacyDocNoteParagraphRuns(
                    lastParagraph,
                    sourceParagraph,
                    keepNoteReferenceMark: false,
                    LegacyDocNoteProjection.Empty,
                    LegacyDocBookmarkProjection.Create(sourceParagraph.Bookmarks, sourceParagraph.StartCharacter, sourceParagraph.EndCharacter));
                ApplyLegacyDocParagraphFormatting(lastParagraph, sourceParagraph.Format, styleSheet);
            }
        }

        private static void AddLegacyDocEndnoteReference(WordParagraph paragraph, LegacyDocEndnote endnote, LegacyDocStyleSheet styleSheet) {
            if (endnote.ParagraphRuns.Count == 0) {
                return;
            }

            WordParagraph reference = paragraph.AddEndNote(endnote.ParagraphRuns[0].Text);
            List<WordParagraph>? noteParagraphs = reference.EndNote!.Paragraphs;
            if (noteParagraphs == null || noteParagraphs.Count == 0) {
                return;
            }

            ReplaceLegacyDocNoteParagraphRuns(
                noteParagraphs[0],
                endnote.ParagraphRuns[0],
                keepNoteReferenceMark: true,
                LegacyDocNoteProjection.Empty,
                LegacyDocBookmarkProjection.Create(endnote.ParagraphRuns[0].Bookmarks, endnote.ParagraphRuns[0].StartCharacter, endnote.ParagraphRuns[0].EndCharacter));
            ApplyLegacyDocParagraphFormatting(noteParagraphs[0], endnote.ParagraphRuns[0].Format, styleSheet);
            WordParagraph lastParagraph = noteParagraphs[0];
            for (int index = 1; index < endnote.ParagraphRuns.Count; index++) {
                LegacyDocNoteParagraph sourceParagraph = endnote.ParagraphRuns[index];
                lastParagraph = lastParagraph.AddParagraph(sourceParagraph.Bookmarks.Count == 0 ? sourceParagraph.Text : string.Empty);
                ReplaceLegacyDocNoteParagraphRuns(
                    lastParagraph,
                    sourceParagraph,
                    keepNoteReferenceMark: false,
                    LegacyDocNoteProjection.Empty,
                    LegacyDocBookmarkProjection.Create(sourceParagraph.Bookmarks, sourceParagraph.StartCharacter, sourceParagraph.EndCharacter));
                ApplyLegacyDocParagraphFormatting(lastParagraph, sourceParagraph.Format, styleSheet);
            }
        }

        private static void ReplaceLegacyDocNoteParagraphRuns(WordParagraph target, LegacyDocNoteParagraph source, bool keepNoteReferenceMark, LegacyDocNoteProjection notes) {
            ReplaceLegacyDocNoteParagraphRuns(target, source, keepNoteReferenceMark, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void ReplaceLegacyDocNoteParagraphRuns(WordParagraph target, LegacyDocNoteParagraph source, bool keepNoteReferenceMark, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            foreach (Run run in target._paragraph.Elements<Run>().ToArray()) {
                if (keepNoteReferenceMark && ContainsLegacyDocNoteReferenceMark(run)) {
                    continue;
                }

                run.Remove();
            }

            WordParagraph paragraph = new WordParagraph(target._document, target._paragraph, newRun: false);
            AddLegacyDocRuns(paragraph, source.Runs, notes, bookmarks);
            bookmarks.EmitRemaining(paragraph._paragraph);
        }

        private static void ReplaceLegacyDocParagraphRuns(WordParagraph target, IReadOnlyList<LegacyDocTextRun> sourceRuns, LegacyDocNoteProjection notes) {
            ReplaceLegacyDocParagraphRuns(target, sourceRuns, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void ReplaceLegacyDocParagraphRuns(WordParagraph target, IReadOnlyList<LegacyDocTextRun> sourceRuns, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            foreach (Run run in target._paragraph.Elements<Run>().ToArray()) {
                run.Remove();
            }

            WordParagraph paragraph = new WordParagraph(target._document, target._paragraph, newRun: false);
            AddLegacyDocRuns(paragraph, sourceRuns, notes, bookmarks);
            bookmarks.EmitRemaining(paragraph._paragraph);
        }

        private static bool ContainsLegacyDocNoteReferenceMark(Run run) {
            return run.Elements<FootnoteReferenceMark>().Any()
                || run.Elements<EndnoteReferenceMark>().Any();
        }

        private sealed class LegacyDocNoteProjection {
            private readonly IReadOnlyDictionary<int, LegacyDocFootnote> _footnotesByReferencePosition;
            private readonly IReadOnlyDictionary<int, LegacyDocEndnote> _endnotesByReferencePosition;

            private LegacyDocNoteProjection(
                IReadOnlyDictionary<int, LegacyDocFootnote> footnotesByReferencePosition,
                IReadOnlyDictionary<int, LegacyDocEndnote> endnotesByReferencePosition,
                LegacyDocStyleSheet styleSheet) {
                _footnotesByReferencePosition = footnotesByReferencePosition;
                _endnotesByReferencePosition = endnotesByReferencePosition;
                StyleSheet = styleSheet;
            }

            internal static LegacyDocNoteProjection Create(IReadOnlyList<LegacyDocFootnote> footnotes, IReadOnlyList<LegacyDocEndnote> endnotes, LegacyDocStyleSheet styleSheet) {
                return new LegacyDocNoteProjection(
                    footnotes
                        .GroupBy(footnote => footnote.ReferenceCharacterPosition)
                        .ToDictionary(group => group.Key, group => group.First()),
                    endnotes
                        .GroupBy(endnote => endnote.ReferenceCharacterPosition)
                        .ToDictionary(group => group.Key, group => group.First()),
                    styleSheet);
            }

            internal static LegacyDocNoteProjection Empty { get; } = new LegacyDocNoteProjection(
                new Dictionary<int, LegacyDocFootnote>(),
                new Dictionary<int, LegacyDocEndnote>(),
                LegacyDocStyleSheet.Empty);

            internal LegacyDocStyleSheet StyleSheet { get; }

            internal bool TryGetFootnote(int referenceCharacterPosition, out LegacyDocFootnote? footnote) {
                return _footnotesByReferencePosition.TryGetValue(referenceCharacterPosition, out footnote);
            }

            internal bool TryGetEndnote(int referenceCharacterPosition, out LegacyDocEndnote? endnote) {
                return _endnotesByReferencePosition.TryGetValue(referenceCharacterPosition, out endnote);
            }
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
            return character == '\t'
                || character == LegacyDocSpecialCharacters.TextWrappingBreak
                || character == LegacyDocSpecialCharacters.PageBreak
                || character == LegacyDocSpecialCharacters.ColumnBreak
                || character == LegacyDocFootnoteReader.FootnoteReferenceCharacter;
        }

        private static bool IsLegacyDocFieldResultSpecialRunCharacter(char character) {
            return character == '\t'
                || character == LegacyDocSpecialCharacters.SoftHyphen
                || character == LegacyDocSpecialCharacters.NoBreakHyphen
                || character == LegacyDocSpecialCharacters.TextWrappingBreak
                || character == LegacyDocSpecialCharacters.PageBreak
                || character == LegacyDocSpecialCharacters.ColumnBreak;
        }

        private static bool IsLegacyDocHyperlinkSpecialRunCharacter(char character) {
            return character == '\t'
                || character == LegacyDocSpecialCharacters.TextWrappingBreak
                || character == LegacyDocSpecialCharacters.PageBreak
                || character == LegacyDocSpecialCharacters.ColumnBreak;
        }

        private static BreakValues? GetLegacyDocBreakType(char character) {
            if (character == LegacyDocSpecialCharacters.PageBreak) {
                return BreakValues.Page;
            }

            if (character == LegacyDocSpecialCharacters.ColumnBreak) {
                return BreakValues.Column;
            }

            return null;
        }

        private static LegacyDocTextRun CreateLegacyDocRunWithoutHyperlink(LegacyDocTextRun source) {
            return new LegacyDocTextRun(
                source.Text,
                source.Bold,
                source.Italic,
                source.Strike,
                source.DoubleStrike,
                source.Outline,
                source.Shadow,
                source.Emboss,
                source.Imprint,
                source.Hidden,
                source.NoProof,
                source.Caps,
                source.VerticalPosition,
                source.Underline,
                source.Highlight,
                source.FontSizeHalfPoints,
                source.ColorHex,
                source.FontFamily,
                source.CharacterPositions,
                fieldKind: source.FieldKind,
                fieldInstruction: source.FieldInstruction,
                specified: source.Specified,
                characterSpacingTwips: source.CharacterSpacingTwips,
                language: source.Language,
                eastAsiaLanguage: source.EastAsiaLanguage);
        }

        private static void AddLegacyDocPageNumber(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var run = new Run(new DocumentFormat.OpenXml.Wordprocessing.PageNumber());
            paragraph._paragraph.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AddLegacyDocNumberOfPages(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var simpleField = new SimpleField { Instruction = " NUMPAGES  " };
            AppendLegacyDocFieldResultContent(simpleField, paragraph, legacyRun, string.IsNullOrEmpty(legacyRun.Text) ? "1" : legacyRun.Text);
            paragraph._paragraph.Append(simpleField);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AddLegacyDocStaticDisplayField(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunCharacterPosition(legacyRun, 0));
            var simpleField = new SimpleField { Instruction = string.IsNullOrWhiteSpace(legacyRun.FieldInstruction) ? GetLegacyDocStaticFieldInstruction(legacyRun.FieldKind) : legacyRun.FieldInstruction };
            AppendLegacyDocFieldResultContent(simpleField, paragraph, legacyRun, legacyRun.Text);
            paragraph._paragraph.Append(simpleField);
            bookmarks.EmitAt(paragraph._paragraph, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AppendLegacyDocFieldResultContent(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, string resultText) {
            int segmentStart = 0;
            for (int index = 0; index < resultText.Length; index++) {
                char character = resultText[index];
                if (!IsLegacyDocFieldResultSpecialRunCharacter(character)) {
                    continue;
                }

                AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, resultText, segmentStart, index - segmentStart);
                AppendLegacyDocFieldSpecialRun(simpleField, paragraph, legacyRun, character);
                segmentStart = index + 1;
            }

            AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, resultText, segmentStart, resultText.Length - segmentStart);
            if (!simpleField.HasChildren) {
                AppendLegacyDocFieldTextSegment(simpleField, paragraph, legacyRun, "1", 0, 1);
            }
        }

        private static void AppendLegacyDocFieldTextSegment(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, string text, int startIndex, int length) {
            if (length <= 0) {
                return;
            }

            var run = new Run(new Text(text.Substring(startIndex, length)) {
                Space = SpaceProcessingModeValues.Preserve
            });
            simpleField.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static void AppendLegacyDocFieldSpecialRun(SimpleField simpleField, WordParagraph paragraph, LegacyDocTextRun legacyRun, char character) {
            var run = new Run();
            if (character == '\t') {
                run.Append(new TabChar());
            } else if (character == LegacyDocSpecialCharacters.SoftHyphen) {
                run.Append(new SoftHyphen());
            } else if (character == LegacyDocSpecialCharacters.NoBreakHyphen) {
                run.Append(new NoBreakHyphen());
            } else {
                run.Append(character == LegacyDocSpecialCharacters.TextWrappingBreak
                    ? new Break()
                    : new Break { Type = GetLegacyDocBreakType(character) });
            }

            simpleField.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static string GetLegacyDocStaticFieldInstruction(LegacyDocFieldKind fieldKind) {
            return fieldKind switch {
                LegacyDocFieldKind.Time => " TIME  ",
                LegacyDocFieldKind.CreateDate => " CREATEDATE  ",
                LegacyDocFieldKind.SaveDate => " SAVEDATE  ",
                LegacyDocFieldKind.PrintDate => " PRINTDATE  ",
                LegacyDocFieldKind.DocumentProperty => throw new NotSupportedException("Legacy DOC document-property field projection requires the source field instruction."),
                _ => " DATE  "
            };
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

        private static void AddLegacyDocHyperlinkRunContent(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes) {
            AddLegacyDocHyperlinkRunsContent(paragraph, new[] { legacyRun }, 0, 1, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void AddLegacyDocHyperlinkRunContent(WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            AddLegacyDocHyperlinkRunsContent(paragraph, new[] { legacyRun }, 0, 1, notes, bookmarks);
        }

        private static void AddLegacyDocHyperlinkRunsContent(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> legacyRuns, int startIndex, int count, LegacyDocNoteProjection notes) {
            AddLegacyDocHyperlinkRunsContent(paragraph, legacyRuns, startIndex, count, notes, LegacyDocBookmarkProjection.Empty);
        }

        private static void AddLegacyDocHyperlinkRunsContent(WordParagraph paragraph, IReadOnlyList<LegacyDocTextRun> legacyRuns, int startIndex, int count, LegacyDocNoteProjection notes, LegacyDocBookmarkProjection bookmarks) {
            LegacyDocTextRun firstRun = legacyRuns[startIndex];
            LegacyDocHyperlinkTarget target = firstRun.HyperlinkTarget;
            if (target.Uri == null && target.Anchor == null) {
                for (int index = startIndex; index < startIndex + count; index++) {
                    AddLegacyDocRunContent(paragraph, CreateLegacyDocRunWithoutHyperlink(legacyRuns[index]), notes, bookmarks);
                }

                return;
            }

            Hyperlink hyperlink;
            if (target.Uri != null && Uri.TryCreate(target.Uri, UriKind.Absolute, out Uri? uri)) {
                OpenXmlPart relationshipOwner = ResolveLegacyDocHyperlinkRelationshipPart(paragraph)
                    ?? throw new InvalidOperationException("Legacy DOC hyperlink projection requires a document relationship owner.");
                HyperlinkRelationship relationship = relationshipOwner.AddHyperlinkRelationship(uri, true);
                hyperlink = new Hyperlink {
                    Id = relationship.Id,
                    History = true
                };
            } else if (target.Anchor != null) {
                hyperlink = new Hyperlink {
                    Anchor = target.Anchor,
                    History = true
                };
            } else {
                for (int index = startIndex; index < startIndex + count; index++) {
                    AddLegacyDocRunContent(paragraph, CreateLegacyDocRunWithoutHyperlink(legacyRuns[index]), notes, bookmarks);
                }

                return;
            }

            for (int index = startIndex; index < startIndex + count; index++) {
                AppendLegacyDocHyperlinkRunContent(hyperlink, paragraph, legacyRuns[index], bookmarks);
            }

            paragraph._paragraph.Append(hyperlink);
            paragraph._hyperlink = hyperlink;
        }

        private static void AppendLegacyDocHyperlinkRunContent(Hyperlink hyperlink, WordParagraph paragraph, LegacyDocTextRun legacyRun) {
            AppendLegacyDocHyperlinkRunContent(hyperlink, paragraph, legacyRun, LegacyDocBookmarkProjection.Empty);
        }

        private static void AppendLegacyDocHyperlinkRunContent(Hyperlink hyperlink, WordParagraph paragraph, LegacyDocTextRun legacyRun, LegacyDocBookmarkProjection bookmarks) {
            string text = legacyRun.Text;
            int segmentStart = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (!IsLegacyDocHyperlinkSpecialRunCharacter(character)) {
                    int? markerPosition = GetLegacyDocRunCharacterPosition(legacyRun, index);
                    if (!bookmarks.HasMarkers(markerPosition)) {
                        continue;
                    }

                    AppendLegacyDocHyperlinkTextRun(hyperlink, paragraph, legacyRun, segmentStart, index - segmentStart);
                    bookmarks.EmitAt(hyperlink, markerPosition);
                    segmentStart = index;
                    continue;
                }

                AppendLegacyDocHyperlinkTextRun(hyperlink, paragraph, legacyRun, segmentStart, index - segmentStart);
                bookmarks.EmitAt(hyperlink, GetLegacyDocRunCharacterPosition(legacyRun, index));
                if (character == '\t') {
                    AppendLegacyDocHyperlinkSpecialRun(hyperlink, paragraph, legacyRun, new TabChar());
                } else {
                    BreakValues? breakType = GetLegacyDocBreakType(character);
                    AppendLegacyDocHyperlinkSpecialRun(hyperlink, paragraph, legacyRun, breakType == null ? new Break() : new Break { Type = breakType });
                }

                segmentStart = index + 1;
            }

            AppendLegacyDocHyperlinkTextRun(hyperlink, paragraph, legacyRun, segmentStart, text.Length - segmentStart);
            bookmarks.EmitAt(hyperlink, GetLegacyDocRunEndCharacterPosition(legacyRun));
        }

        private static void AppendLegacyDocHyperlinkTextRun(Hyperlink hyperlink, WordParagraph paragraph, LegacyDocTextRun legacyRun, int startIndex, int length) {
            if (length <= 0) {
                return;
            }

            var run = new Run(new Text(legacyRun.Text.Substring(startIndex, length)) {
                Space = SpaceProcessingModeValues.Preserve
            });
            hyperlink.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static void AppendLegacyDocHyperlinkSpecialRun(Hyperlink hyperlink, WordParagraph paragraph, LegacyDocTextRun legacyRun, OpenXmlElement element) {
            var run = new Run();
            run.Append(element);
            hyperlink.Append(run);
            ApplyLegacyDocRunFormatting(new WordParagraph(paragraph._document, paragraph._paragraph, run), legacyRun);
        }

        private static OpenXmlPart? ResolveLegacyDocHyperlinkRelationshipPart(WordParagraph paragraph) {
            OpenXmlElement? parent = paragraph._paragraph.Parent;
            while (parent != null
                && parent is not Body
                && parent is not DocumentFormat.OpenXml.Wordprocessing.Header
                && parent is not DocumentFormat.OpenXml.Wordprocessing.Footer
                && parent is not Footnote
                && parent is not Endnote) {
                parent = parent.Parent;
            }

            HeaderPart? headerPart = (parent as DocumentFormat.OpenXml.Wordprocessing.Header)?.HeaderPart;
            if (headerPart != null) {
                return headerPart;
            }

            FooterPart? footerPart = (parent as DocumentFormat.OpenXml.Wordprocessing.Footer)?.FooterPart;
            if (footerPart != null) {
                return footerPart;
            }

            MainDocumentPart? mainPart = paragraph._document._wordprocessingDocument?.MainDocumentPart;
            if (parent is Footnote) {
                return mainPart?.FootnotesPart;
            }

            if (parent is Endnote) {
                return mainPart?.EndnotesPart;
            }

            return mainPart;
        }

        private static void ApplyLegacyDocParagraphFormatting(WordParagraph paragraph, LegacyDocParagraphFormat paragraphFormat, LegacyDocStyleSheet styleSheet) {
            if (paragraphFormat.StyleIndex != null) {
                ApplyLegacyDocParagraphStyle(paragraph, paragraphFormat.StyleIndex.Value, styleSheet);
            }

            if (paragraphFormat.ParagraphMarkFormat != null) {
                ApplyLegacyDocParagraphMarkRunFormatting(paragraph, paragraphFormat.ParagraphMarkFormat.Value);
            }

            if (paragraphFormat.NumberingListIndex != null) {
                ApplyLegacyDocParagraphNumbering(paragraph, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0);
            }

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                paragraph.ParagraphAlignment = alignment;
            }

            if (paragraphFormat.VerticalCharacterAlignment != null && TryMapVerticalCharacterAlignment(paragraphFormat.VerticalCharacterAlignment.Value, out VerticalTextAlignmentValues verticalCharacterAlignment)) {
                paragraph.VerticalCharacterAlignmentOnLine = verticalCharacterAlignment;
            }

            if (paragraphFormat.OutlineLevel != null) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new OutlineLevel { Val = paragraphFormat.OutlineLevel.Value });
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

            if (paragraphFormat.SuppressLineNumbers == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new SuppressLineNumbers());
            }

            if (paragraphFormat.SuppressAutoHyphens == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new SuppressAutoHyphens());
            }

            if (paragraphFormat.ContextualSpacing == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new ContextualSpacing());
            }

            if (paragraphFormat.MirrorIndents == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new MirrorIndents());
            }

            if (paragraphFormat.Kinsoku == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new Kinsoku());
            }

            if (paragraphFormat.WordWrap == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new WordWrap());
            }

            if (paragraphFormat.OverflowPunctuation == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new OverflowPunctuation());
            }

            if (paragraphFormat.TopLinePunctuation == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new TopLinePunctuation());
            }

            if (paragraphFormat.AutoSpaceDE == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new AutoSpaceDE());
            }

            if (paragraphFormat.AutoSpaceDN == true) {
                EnsureLegacyDocParagraphProperties(paragraph).Append(new AutoSpaceDN());
            }

            if (paragraphFormat.Bidirectional == true) {
                paragraph.BiDi = true;
            }

            if (paragraphFormat.ParagraphShading != null && !string.IsNullOrEmpty(paragraphFormat.ParagraphShading.Value.FillColorHex)) {
                paragraph.ShadingFillColorHex = paragraphFormat.ParagraphShading.Value.FillColorHex!;
            }

            if (paragraphFormat.ParagraphBorders != null && paragraphFormat.ParagraphBorders.Value.HasAny) {
                ApplyLegacyDocParagraphBorders(paragraph, paragraphFormat.ParagraphBorders.Value);
            }

            foreach (LegacyDocTabStop tabStop in paragraphFormat.TabStops) {
                if (TryMapTabStopAlignment(tabStop.Alignment, out TabStopValues tabAlignment)
                    && TryMapTabStopLeader(tabStop.Leader, out TabStopLeaderCharValues leader)) {
                    paragraph.AddTabStop(tabStop.PositionTwips, tabAlignment, leader);
                }
            }
        }

        private static void ApplyLegacyDocParagraphMarkRunFormatting(WordParagraph paragraph, LegacyDocCharacterFormat characterFormat) {
            StyleRunProperties? styleRunProperties = CreateLegacyDocStyleRunProperties(characterFormat);
            if (styleRunProperties == null) {
                return;
            }

            ParagraphProperties paragraphProperties = EnsureLegacyDocParagraphProperties(paragraph);
            paragraphProperties.RemoveAllChildren<ParagraphMarkRunProperties>();
            var paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            foreach (OpenXmlElement property in styleRunProperties.ChildElements) {
                paragraphMarkRunProperties.Append(property.CloneNode(true));
            }

            paragraphProperties.Append(paragraphMarkRunProperties);
        }

        private static bool TryMapVerticalCharacterAlignment(byte alignment, out VerticalTextAlignmentValues verticalCharacterAlignment) {
            switch (alignment) {
                case 0:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Auto;
                    return true;
                case 1:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Baseline;
                    return true;
                case 2:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Top;
                    return true;
                case 3:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Center;
                    return true;
                case 4:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Bottom;
                    return true;
                default:
                    verticalCharacterAlignment = VerticalTextAlignmentValues.Auto;
                    return false;
            }
        }

        private static void ApplyLegacyDocParagraphNumbering(WordParagraph paragraph, ushort listIndex, byte level) {
            EnsureLegacyDocNumberingDefinition(paragraph._document, listIndex);

            ParagraphProperties paragraphProperties = EnsureLegacyDocParagraphProperties(paragraph);
            NumberingProperties? numberingProperties = paragraphProperties.GetFirstChild<NumberingProperties>();
            if (numberingProperties == null) {
                numberingProperties = new NumberingProperties();
                paragraphProperties.Append(numberingProperties);
            }

            ReplaceLegacyDocNumberingProperties(numberingProperties, listIndex, level);
        }

        private static NumberingProperties CreateLegacyDocNumberingProperties(WordDocument document, ushort listIndex, byte level) {
            EnsureLegacyDocNumberingDefinition(document, listIndex);
            var numberingProperties = new NumberingProperties();
            ReplaceLegacyDocNumberingProperties(numberingProperties, listIndex, level);
            return numberingProperties;
        }

        private static void ReplaceLegacyDocNumberingProperties(NumberingProperties numberingProperties, ushort listIndex, byte level) {
            numberingProperties.RemoveAllChildren<NumberingLevelReference>();
            numberingProperties.RemoveAllChildren<NumberingId>();
            numberingProperties.Append(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = listIndex });
        }

        private static ParagraphProperties EnsureLegacyDocParagraphProperties(WordParagraph paragraph) {
            if (paragraph._paragraph.ParagraphProperties == null) {
                paragraph._paragraph.ParagraphProperties = new ParagraphProperties();
            }

            return paragraph._paragraph.ParagraphProperties!;
        }

        private static void EnsureLegacyDocNumberingDefinition(WordDocument document, int numberId) {
            MainDocumentPart mainPart = document.MainDocumentPartRoot;
            NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            Numbering numbering = numberingPart.Numbering ??= new Numbering();

            if (numbering.Elements<NumberingInstance>().Any(instance => instance.NumberID?.Value == numberId)) {
                return;
            }

            int abstractId = GetNextLegacyDocAbstractNumberingId(numbering);
            numbering.Append(CreateLegacyDocDecimalAbstractNumbering(abstractId));
            numbering.Append(new NumberingInstance(new AbstractNumId { Val = abstractId }) { NumberID = numberId });
        }

        private static int GetNextLegacyDocAbstractNumberingId(Numbering numbering) {
            int nextId = 0;
            foreach (AbstractNum abstractNum in numbering.Elements<AbstractNum>()) {
                if (abstractNum.AbstractNumberId?.Value != null) {
                    nextId = Math.Max(nextId, abstractNum.AbstractNumberId.Value + 1);
                }
            }

            return nextId;
        }

        private static AbstractNum CreateLegacyDocDecimalAbstractNumbering(int abstractId) {
            var abstractNum = new AbstractNum {
                AbstractNumberId = abstractId
            };
            abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

            for (int level = 0; level <= 8; level++) {
                abstractNum.Append(CreateLegacyDocDecimalNumberingLevel(level));
            }

            return abstractNum;
        }

        private static Level CreateLegacyDocDecimalNumberingLevel(int level) {
            var paragraphProperties = new PreviousParagraphProperties();
            paragraphProperties.Append(new Indentation {
                Left = ((level + 1) * 720).ToString(System.Globalization.CultureInfo.InvariantCulture),
                Hanging = "360"
            });

            var numberingLevel = new Level {
                LevelIndex = level,
                Tentative = level > 0
            };
            numberingLevel.Append(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = MapLegacyDocFallbackNumberingFormat(level) },
                new LevelText { Val = "%" + (level + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + "." },
                new LevelJustification { Val = LevelJustificationValues.Left },
                paragraphProperties);
            return numberingLevel;
        }

        private static NumberFormatValues MapLegacyDocFallbackNumberingFormat(int level) {
            switch (level % 3) {
                case 1:
                    return NumberFormatValues.LowerLetter;
                case 2:
                    return NumberFormatValues.LowerRoman;
                default:
                    return NumberFormatValues.Decimal;
            }
        }

        private static void ApplyLegacyDocParagraphBorders(WordParagraph paragraph, LegacyDocParagraphBorders borders) {
            ApplyLegacyDocParagraphBorder(
                borders.Top,
                style => paragraph.Borders.TopStyle = style,
                color => paragraph.Borders.TopColorHex = color,
                size => paragraph.Borders.TopSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => paragraph.Borders.TopSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocParagraphBorder(
                borders.Left,
                style => paragraph.Borders.LeftStyle = style,
                color => paragraph.Borders.LeftColorHex = color,
                size => paragraph.Borders.LeftSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => paragraph.Borders.LeftSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocParagraphBorder(
                borders.Bottom,
                style => paragraph.Borders.BottomStyle = style,
                color => paragraph.Borders.BottomColorHex = color,
                size => paragraph.Borders.BottomSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => paragraph.Borders.BottomSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);
            ApplyLegacyDocParagraphBorder(
                borders.Right,
                style => paragraph.Borders.RightStyle = style,
                color => paragraph.Borders.RightColorHex = color,
                size => paragraph.Borders.RightSize = (DocumentFormat.OpenXml.UInt32Value)(uint)size,
                space => paragraph.Borders.RightSpace = (DocumentFormat.OpenXml.UInt32Value)(uint)space);

            if (borders.Between.HasAny) {
                ParagraphBorders paragraphBorders = GetOrCreateParagraphBorders(paragraph);
                paragraphBorders.BetweenBorder = CreateLegacyDocParagraphBorder<BetweenBorder>(borders.Between);
            }
        }

        private static void ApplyLegacyDocParagraphBorder(
            LegacyDocParagraphBorder border,
            Action<BorderValues> setStyle,
            Action<string> setColor,
            Action<int> setSize,
            Action<int> setSpace) {
            BorderValues? style = MapLegacyDocParagraphBorderStyle(border.Style);
            if (style == null) {
                return;
            }

            setStyle(style.Value);
            if (!string.IsNullOrEmpty(border.ColorHex)) {
                setColor(border.ColorHex!);
            }

            if (border.SizeEighthPoints > 0) {
                setSize(border.SizeEighthPoints);
            }

            if (border.SpacePoints > 0) {
                setSpace(border.SpacePoints);
            }
        }

        private static ParagraphBorders GetOrCreateParagraphBorders(WordParagraph paragraph) {
            ParagraphBorders? borders = paragraph._paragraphProperties?.GetFirstChild<ParagraphBorders>();
            if (borders == null) {
                borders = new ParagraphBorders();
                paragraph._paragraphProperties!.Append(borders);
            }

            return borders;
        }

        private static TBorder CreateLegacyDocParagraphBorder<TBorder>(LegacyDocParagraphBorder border) where TBorder : BorderType, new() {
            BorderValues? style = MapLegacyDocParagraphBorderStyle(border.Style);
            var paragraphBorder = new TBorder {
                Val = style.GetValueOrDefault(),
                Size = (DocumentFormat.OpenXml.UInt32Value)(uint)border.SizeEighthPoints,
                Space = (DocumentFormat.OpenXml.UInt32Value)(uint)border.SpacePoints
            };
            if (!string.IsNullOrEmpty(border.ColorHex)) {
                paragraphBorder.Color = border.ColorHex!;
            }

            return paragraphBorder;
        }

        private static BorderValues? MapLegacyDocParagraphBorderStyle(LegacyDocParagraphBorderStyle style) {
            switch (style) {
                case LegacyDocParagraphBorderStyle.Single:
                    return BorderValues.Single;
                case LegacyDocParagraphBorderStyle.Double:
                    return BorderValues.Double;
                case LegacyDocParagraphBorderStyle.Dotted:
                    return BorderValues.Dotted;
                case LegacyDocParagraphBorderStyle.Dashed:
                    return BorderValues.Dashed;
                default:
                    return null;
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
                if (legacyStyle.BuiltInStyle != null) {
                    string builtInStyleId = legacyStyle.BuiltInStyle.Value.ToStringStyle();
                    Style builtInStyle = GetOrCreateLegacyDocBuiltInStyle(styles, builtInStyleId, legacyStyle.Name);
                    MergeLegacyDocBuiltInStyleFormatting(document, builtInStyle, legacyStyle, styleSheet);
                    AddOrReplaceLegacyDocNextStyle(builtInStyle, legacyStyle, styleSheet);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(legacyStyle.StyleId)) {
                    continue;
                }

                if (!existingStyleIds.Add(legacyStyle.StyleId!)) {
                    continue;
                }

                var customStyle = new Style { Type = StyleValues.Paragraph, StyleId = legacyStyle.StyleId, CustomStyle = true };
                customStyle.Append(new StyleName { Val = legacyStyle.Name });
                customStyle.Append(new BasedOn { Val = ResolveLegacyDocBasedOnStyleId(legacyStyle, styleSheet) });
                string? nextStyleId = ResolveLegacyDocNextStyleId(legacyStyle, styleSheet);
                if (!string.IsNullOrWhiteSpace(nextStyleId)) {
                    customStyle.Append(new NextParagraphStyle { Val = nextStyleId });
                }

                StyleParagraphProperties? styleParagraphProperties = CreateLegacyDocStyleParagraphProperties(document, legacyStyle.ParagraphFormat);
                if (styleParagraphProperties != null) {
                    customStyle.Append(styleParagraphProperties);
                }

                StyleRunProperties? styleRunProperties = CreateLegacyDocStyleRunProperties(legacyStyle.CharacterFormat);
                if (styleRunProperties != null) {
                    customStyle.Append(styleRunProperties);
                }

                styles.Append(customStyle);
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

        private static void AddOrReplaceLegacyDocNextStyle(Style style, LegacyDocParagraphStyle legacyStyle, LegacyDocStyleSheet styleSheet) {
            string? nextStyleId = ResolveLegacyDocNextStyleId(legacyStyle, styleSheet);
            if (!string.IsNullOrWhiteSpace(nextStyleId)) {
                ReplaceStyleProperty(style, new NextParagraphStyle { Val = nextStyleId });
            }
        }

        private static string? ResolveLegacyDocNextStyleId(LegacyDocParagraphStyle legacyStyle, LegacyDocStyleSheet styleSheet) {
            if (legacyStyle.NextStyleIndex == null || legacyStyle.NextStyleIndex.Value == legacyStyle.Index) {
                return null;
            }

            if (styleSheet.TryGetParagraphStyle(legacyStyle.NextStyleIndex.Value, out LegacyDocParagraphStyle nextStyle)) {
                if (nextStyle.BuiltInStyle != null) {
                    return nextStyle.BuiltInStyle.Value.ToStringStyle();
                }

                if (!string.IsNullOrWhiteSpace(nextStyle.StyleId)) {
                    return nextStyle.StyleId!;
                }
            }

            if (TryMapBuiltInParagraphStyle(legacyStyle.NextStyleIndex.Value, out WordParagraphStyles builtInStyle)) {
                return builtInStyle.ToStringStyle();
            }

            return null;
        }

        private static StyleParagraphProperties? CreateLegacyDocStyleParagraphProperties(WordDocument document, LegacyDocParagraphFormat paragraphFormat) {
            var properties = new StyleParagraphProperties();
            bool hasProperties = false;

            if (paragraphFormat.Alignment != null && TryMapParagraphAlignment(paragraphFormat.Alignment.Value, out JustificationValues alignment)) {
                properties.Append(new Justification { Val = alignment });
                hasProperties = true;
            }

            if (paragraphFormat.NumberingListIndex != null) {
                properties.Append(CreateLegacyDocNumberingProperties(document, paragraphFormat.NumberingListIndex.Value, paragraphFormat.NumberingLevel ?? 0));
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

            Tabs? tabs = CreateLegacyDocTabs(paragraphFormat.TabStops);
            if (tabs != null) {
                properties.Append(tabs);
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

            if (paragraphFormat.SuppressLineNumbers == true) {
                properties.Append(new SuppressLineNumbers());
                hasProperties = true;
            }

            if (paragraphFormat.SuppressAutoHyphens == true) {
                properties.Append(new SuppressAutoHyphens());
                hasProperties = true;
            }

            if (paragraphFormat.ContextualSpacing == true) {
                properties.Append(new ContextualSpacing());
                hasProperties = true;
            }

            if (paragraphFormat.MirrorIndents == true) {
                properties.Append(new MirrorIndents());
                hasProperties = true;
            }

            if (paragraphFormat.Kinsoku == true) {
                properties.Append(new Kinsoku());
                hasProperties = true;
            }

            if (paragraphFormat.WordWrap == true) {
                properties.Append(new WordWrap());
                hasProperties = true;
            }

            if (paragraphFormat.OverflowPunctuation == true) {
                properties.Append(new OverflowPunctuation());
                hasProperties = true;
            }

            if (paragraphFormat.TopLinePunctuation == true) {
                properties.Append(new TopLinePunctuation());
                hasProperties = true;
            }

            if (paragraphFormat.AutoSpaceDE == true) {
                properties.Append(new AutoSpaceDE());
                hasProperties = true;
            }

            if (paragraphFormat.AutoSpaceDN == true) {
                properties.Append(new AutoSpaceDN());
                hasProperties = true;
            }

            if (paragraphFormat.Bidirectional == true) {
                properties.Append(new BiDi());
                hasProperties = true;
            }

            if (paragraphFormat.VerticalCharacterAlignment != null && TryMapVerticalCharacterAlignment(paragraphFormat.VerticalCharacterAlignment.Value, out VerticalTextAlignmentValues verticalCharacterAlignment)) {
                properties.Append(new TextAlignment { Val = verticalCharacterAlignment });
                hasProperties = true;
            }

            if (paragraphFormat.OutlineLevel != null) {
                properties.Append(new OutlineLevel { Val = paragraphFormat.OutlineLevel.Value });
                hasProperties = true;
            }

            if (paragraphFormat.ParagraphShading != null && !string.IsNullOrEmpty(paragraphFormat.ParagraphShading.Value.FillColorHex)) {
                properties.Append(new Shading {
                    Val = ShadingPatternValues.Clear,
                    Color = "auto",
                    Fill = paragraphFormat.ParagraphShading.Value.FillColorHex!
                });
                hasProperties = true;
            }

            if (paragraphFormat.ParagraphBorders != null && paragraphFormat.ParagraphBorders.Value.HasAny) {
                properties.Append(CreateLegacyDocStyleParagraphBorders(paragraphFormat.ParagraphBorders.Value));
                hasProperties = true;
            }

            return hasProperties ? properties : null;
        }

        private static ParagraphBorders CreateLegacyDocStyleParagraphBorders(LegacyDocParagraphBorders borders) {
            var paragraphBorders = new ParagraphBorders();
            if (borders.Top.HasAny) {
                paragraphBorders.TopBorder = CreateLegacyDocParagraphBorder<TopBorder>(borders.Top);
            }

            if (borders.Left.HasAny) {
                paragraphBorders.LeftBorder = CreateLegacyDocParagraphBorder<LeftBorder>(borders.Left);
            }

            if (borders.Bottom.HasAny) {
                paragraphBorders.BottomBorder = CreateLegacyDocParagraphBorder<BottomBorder>(borders.Bottom);
            }

            if (borders.Right.HasAny) {
                paragraphBorders.RightBorder = CreateLegacyDocParagraphBorder<RightBorder>(borders.Right);
            }

            if (borders.Between.HasAny) {
                paragraphBorders.BetweenBorder = CreateLegacyDocParagraphBorder<BetweenBorder>(borders.Between);
            }

            return paragraphBorders;
        }

        private static Tabs? CreateLegacyDocTabs(IReadOnlyList<LegacyDocTabStop> tabStops) {
            if (tabStops.Count == 0) {
                return null;
            }

            var tabs = new Tabs();
            foreach (LegacyDocTabStop tabStop in tabStops) {
                if (TryMapTabStopAlignment(tabStop.Alignment, out TabStopValues alignment)
                    && TryMapTabStopLeader(tabStop.Leader, out TabStopLeaderCharValues leader)) {
                    tabs.Append(new TabStop {
                        Val = alignment,
                        Leader = leader,
                        Position = tabStop.PositionTwips
                    });
                }
            }

            return tabs.HasChildren ? tabs : null;
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

            if (!string.IsNullOrEmpty(characterFormat.Language) || !string.IsNullOrEmpty(characterFormat.EastAsiaLanguage)) {
                properties.Append(new Languages {
                    Val = characterFormat.Language,
                    EastAsia = characterFormat.EastAsiaLanguage
                });
                hasProperties = true;
            }

            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Bold>(properties, characterFormat.Bold, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<BoldComplexScript>(properties, characterFormat.Bold, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Italic>(properties, characterFormat.Italic, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<ItalicComplexScript>(properties, characterFormat.Italic, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Strike>(properties, characterFormat.Strike, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<DoubleStrike>(properties, characterFormat.DoubleStrike, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Outline>(properties, characterFormat.Outline, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Shadow>(properties, characterFormat.Shadow, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Emboss>(properties, characterFormat.Emboss, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Imprint>(properties, characterFormat.Imprint, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Vanish>(properties, characterFormat.Hidden, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<NoProof>(properties, characterFormat.NoProof, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<Caps>(properties, characterFormat.Caps == LegacyDocCapsKind.Caps, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            hasProperties |= AppendLegacyDocStyleRunOnOffProperty<SmallCaps>(properties, characterFormat.Caps == LegacyDocCapsKind.SmallCaps, characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));

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

            if (characterFormat.CharacterSpacingTwips != null || characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                properties.Append(new Spacing { Val = characterFormat.CharacterSpacingTwips ?? 0 });
                hasProperties = true;
            }

            if (characterFormat.Highlight != null && TryMapHighlight(characterFormat.Highlight.Value, out HighlightColorValues highlight)) {
                properties.Append(new Highlight { Val = highlight });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) {
                properties.Append(new Highlight { Val = HighlightColorValues.None });
                hasProperties = true;
            }

            if (characterFormat.Underline != null && TryMapUnderline(characterFormat.Underline.Value, out UnderlineValues underline)) {
                properties.Append(new Underline { Val = underline });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) {
                properties.Append(new Underline { Val = UnderlineValues.None });
                hasProperties = true;
            }

            if (characterFormat.VerticalPosition != null && TryMapVerticalPosition(characterFormat.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                properties.Append(new VerticalTextAlignment { Val = verticalPosition });
                hasProperties = true;
            } else if (characterFormat.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) {
                properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Baseline });
                hasProperties = true;
            }

            return hasProperties ? properties : null;
        }

        private static bool AppendLegacyDocStyleRunOnOffProperty<T>(StyleRunProperties properties, bool enabled, bool specified) where T : OnOffType, new() {
            if (!enabled && !specified) {
                return false;
            }

            var property = new T();
            if (!enabled) {
                property.Val = false;
            }

            properties.Append(property);
            return true;
        }

        private static void ApplyLegacyDocRunOnOffProperty<T>(WordParagraph run, bool enabled, bool specified) where T : OnOffType, new() {
            if (!enabled && !specified) {
                return;
            }

            RunProperties runProperties = run._runProperties ?? new RunProperties();
            run._runProperties = runProperties;
            runProperties.RemoveAllChildren<T>();

            var property = new T();
            if (!enabled) {
                property.Val = false;
            }

            runProperties.Append(property);
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
            ApplyLegacyDocRunOnOffProperty<Bold>(run, legacyRun.Bold, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ApplyLegacyDocRunOnOffProperty<BoldComplexScript>(run, legacyRun.Bold, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Bold));
            ApplyLegacyDocRunOnOffProperty<Italic>(run, legacyRun.Italic, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ApplyLegacyDocRunOnOffProperty<ItalicComplexScript>(run, legacyRun.Italic, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Italic));
            ApplyLegacyDocRunOnOffProperty<Strike>(run, legacyRun.Strike, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Strike));
            ApplyLegacyDocRunOnOffProperty<DoubleStrike>(run, legacyRun.DoubleStrike, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.DoubleStrike));
            ApplyLegacyDocRunOnOffProperty<Outline>(run, legacyRun.Outline, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Outline));
            ApplyLegacyDocRunOnOffProperty<Shadow>(run, legacyRun.Shadow, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Shadow));
            ApplyLegacyDocRunOnOffProperty<Emboss>(run, legacyRun.Emboss, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Emboss));
            ApplyLegacyDocRunOnOffProperty<Imprint>(run, legacyRun.Imprint, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Imprint));
            ApplyLegacyDocRunOnOffProperty<Vanish>(run, legacyRun.Hidden, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Hidden));
            ApplyLegacyDocRunOnOffProperty<NoProof>(run, legacyRun.NoProof, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.NoProof));
            ApplyLegacyDocRunOnOffProperty<Caps>(run, legacyRun.Caps == LegacyDocCapsKind.Caps, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Caps));
            ApplyLegacyDocRunOnOffProperty<SmallCaps>(run, legacyRun.Caps == LegacyDocCapsKind.SmallCaps, legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.SmallCaps));

            if (legacyRun.VerticalPosition != null && TryMapVerticalPosition(legacyRun.VerticalPosition.Value, out VerticalPositionValues verticalPosition)) {
                run.VerticalTextAlignment = verticalPosition;
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.VerticalPosition)) {
                run.VerticalTextAlignment = VerticalPositionValues.Baseline;
            }

            if (legacyRun.Underline != null && TryMapUnderline(legacyRun.Underline.Value, out UnderlineValues underline)) {
                run.Underline = underline;
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Underline)) {
                run.Underline = UnderlineValues.None;
            }

            if (legacyRun.Highlight != null && TryMapHighlight(legacyRun.Highlight.Value, out HighlightColorValues highlight)) {
                run.Highlight = highlight;
            } else if (legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.Highlight)) {
                run.Highlight = HighlightColorValues.None;
            }

            if (legacyRun.FontSizeHalfPoints != null) {
                string fontSize = legacyRun.FontSizeHalfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.FontSize = new FontSize {
                    Val = fontSize
                };
                runProperties.FontSizeComplexScript = new FontSizeComplexScript {
                    Val = fontSize
                };
            }

            if (!string.IsNullOrEmpty(legacyRun.ColorHex)) {
                run.ColorHex = legacyRun.ColorHex!;
            }

            if (!string.IsNullOrEmpty(legacyRun.FontFamily)) {
                run.SetFontFamily(legacyRun.FontFamily!);
            }

            if (legacyRun.CharacterSpacingTwips != null || legacyRun.IsSpecified(LegacyDocCharacterFormatProperties.CharacterSpacing)) {
                run.Spacing = legacyRun.CharacterSpacingTwips ?? 0;
            }

            if (!string.IsNullOrEmpty(legacyRun.Language) || !string.IsNullOrEmpty(legacyRun.EastAsiaLanguage)) {
                RunProperties runProperties = run._runProperties ?? new RunProperties();
                run._runProperties = runProperties;
                runProperties.Languages = new Languages {
                    Val = legacyRun.Language,
                    EastAsia = legacyRun.EastAsiaLanguage
                };
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

        private static bool TryMapTabStopAlignment(LegacyDocTabStopAlignment alignment, out TabStopValues value) {
            switch (alignment) {
                case LegacyDocTabStopAlignment.Left:
                    value = TabStopValues.Left;
                    return true;
                case LegacyDocTabStopAlignment.Center:
                    value = TabStopValues.Center;
                    return true;
                case LegacyDocTabStopAlignment.Right:
                    value = TabStopValues.Right;
                    return true;
                case LegacyDocTabStopAlignment.Decimal:
                    value = TabStopValues.Decimal;
                    return true;
                case LegacyDocTabStopAlignment.Bar:
                    value = TabStopValues.Bar;
                    return true;
                case LegacyDocTabStopAlignment.Clear:
                    value = TabStopValues.Clear;
                    return true;
                default:
                    value = TabStopValues.Left;
                    return false;
            }
        }

        private static bool TryMapTabStopLeader(LegacyDocTabStopLeader leader, out TabStopLeaderCharValues value) {
            switch (leader) {
                case LegacyDocTabStopLeader.None:
                    value = TabStopLeaderCharValues.None;
                    return true;
                case LegacyDocTabStopLeader.Dot:
                    value = TabStopLeaderCharValues.Dot;
                    return true;
                case LegacyDocTabStopLeader.Hyphen:
                    value = TabStopLeaderCharValues.Hyphen;
                    return true;
                case LegacyDocTabStopLeader.Underscore:
                    value = TabStopLeaderCharValues.Underscore;
                    return true;
                case LegacyDocTabStopLeader.Heavy:
                    value = TabStopLeaderCharValues.Heavy;
                    return true;
                case LegacyDocTabStopLeader.MiddleDot:
                    value = TabStopLeaderCharValues.MiddleDot;
                    return true;
                default:
                    value = TabStopLeaderCharValues.None;
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
