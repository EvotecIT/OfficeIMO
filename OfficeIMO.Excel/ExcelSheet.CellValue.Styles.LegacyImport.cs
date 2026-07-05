using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Creates or reuses an Open XML cell format for a parsed legacy XLS XF record.
        /// </summary>
        internal uint GetOrCreateLegacyCellFormatStyleIndex(LegacyXlsWorkbook workbook, LegacyXlsCellFormat format) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (format == null) throw new ArgumentNullException(nameof(format));

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            uint styleIndex = 0U;
            WriteLockConditional(() => {
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                CellFormat candidate = CreateLegacyCellFormat(stylesheet, workbook, format);

                styleIndex = AppendOrReuseCellFormat(stylesheet, candidate);
                stylesPart.Stylesheet.Save();
            });

            return styleIndex;
        }

        /// <summary>
        /// Converts a parsed legacy XLS XF record into an Open XML cell format.
        /// </summary>
        internal static CellFormat CreateLegacyCellFormat(Stylesheet stylesheet, LegacyXlsWorkbook workbook, LegacyXlsCellFormat format) {
            if (stylesheet == null) throw new ArgumentNullException(nameof(stylesheet));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (format == null) throw new ArgumentNullException(nameof(format));

            EnsureDefaultStylePrimitives(stylesheet);
            CellFormat candidate = GetBaseCellFormat(stylesheet, 0U);
            ApplyLegacyNumberFormat(stylesheet, candidate, format);
            ApplyLegacyFont(stylesheet, candidate, workbook, format);
            ApplyLegacyFill(stylesheet, candidate, workbook, format);
            ApplyLegacyAlignment(candidate, format);
            ApplyLegacyBorder(stylesheet, candidate, workbook, format);
            ApplyLegacyProtection(candidate, format);
            ApplyLegacyQuotePrefix(candidate, format);
            ApplyLegacyCellStyleExtensions(stylesheet, candidate, workbook, format);
            return candidate;
        }

        /// <summary>
        /// Writes an imported legacy XLS error value as an Open XML error cell.
        /// </summary>
        internal void SetLegacyErrorCellValue(int row, int column, string errorText) {
            SetErrorCellValue(row, column, errorText);
        }

        /// <summary>
        /// Writes a native Excel error value into the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="errorText">The Excel error literal, such as <c>#DIV/0!</c> or <c>#N/A</c>.</param>
        public void CellError(int row, int column, string errorText) {
            using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateForExternalCellMutation(this, row, column);
            SetErrorCellValue(row, column, errorText);
        }

        private void SetErrorCellValue(int row, int column, string errorText) {
            if (string.IsNullOrWhiteSpace(errorText)) {
                throw new ArgumentException("Error text must be provided.", nameof(errorText));
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(errorText);
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
                cell.InlineString = null;
                ClearHeaderCacheForCellMutation(row, column);
            });
        }

        /// <summary>
        /// Applies an existing style index to a worksheet column definition.
        /// </summary>
        internal void SetColumnStyleIndex(int columnIndex, uint styleIndex, bool save = true) {
            if (columnIndex <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be greater than 0.");
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                Columns? columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }

                Column? column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column != null) {
                    column = SplitColumn(columns, column, (uint)columnIndex);
                }

                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }

                column.Style = styleIndex;
                ReorderColumns(columns);
                if (save) {
                    worksheet.Save();
                }
            });
        }

        /// <summary>
        /// Applies an existing style index to a worksheet row definition.
        /// </summary>
        internal void SetRowStyleIndex(int rowIndex, uint styleIndex, bool save = true) {
            if (rowIndex <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                Row row = GetOrCreateRowElement(sheetData, rowIndex);
                row.StyleIndex = styleIndex;
                row.CustomFormat = true;
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Applies an existing style index to a worksheet cell.
        /// </summary>
        internal void SetCellStyleIndex(int row, int column, uint styleIndex, bool save = true) {
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row), "Row index must be greater than 0.");
            }

            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column), "Column index must be greater than 0.");
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.StyleIndex = styleIndex;
                ClearHeaderCacheForCellMutation(row, column);
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Applies a built-in Excel number format id while preserving existing style facets on the target cell.
        /// Used by importers that already resolved a legacy or external format to an Excel built-in id.
        /// </summary>
        internal void FormatCellBuiltInNumberFormat(int row, int column, uint builtInFormatId) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => ApplyBuiltInNumberFormat(row, column, builtInFormatId));
        }

        /// <summary>
        /// Applies only the supplied font facets while preserving existing cell style facets.
        /// Null values leave the corresponding font facet unchanged.
        /// </summary>
        internal void FormatCellFont(
            int row,
            int column,
            string? fontName,
            double? fontSize,
            string? fontColor,
            bool? bold,
            bool? italic,
            bool? underline,
            bool? strike = null,
            UnderlineValues? underlineStyle = null,
            VerticalAlignmentRunValues? verticalTextAlignment = null,
            byte? fontFamily = null,
            byte? fontCharacterSet = null,
            bool? outline = null,
            bool? shadow = null,
            bool? condense = null,
            bool? extend = null) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                uint baseIndex = cell.StyleIndex?.Value ?? 0U;
                var baseFormat = GetBaseCellFormat(stylesheet, baseIndex);
                var fontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(baseFormat.FontId), font => {
                    if (!string.IsNullOrWhiteSpace(fontName)) {
                        SetFontName(font, fontName!);
                    }

                    if (fontSize.HasValue && fontSize.Value > 0) {
                        SetFontSize(font, fontSize.Value);
                    }

                    if (!string.IsNullOrWhiteSpace(fontColor)) {
                        SetFontColor(font, fontColor!);
                    }

                    if (bold.HasValue) {
                        SetBold(font, bold.Value);
                    }

                    if (italic.HasValue) {
                        SetItalic(font, italic.Value);
                    }

                    if (underlineStyle.HasValue) {
                        SetUnderline(font, underlineStyle.Value);
                    } else if (underline.HasValue) {
                        SetUnderline(font, underline.Value);
                    }

                    if (strike.HasValue) {
                        SetStrike(font, strike.Value);
                    }

                    if (outline.HasValue) {
                        SetOutline(font, outline.Value);
                    }

                    if (shadow.HasValue) {
                        SetShadow(font, shadow.Value);
                    }

                    if (condense.HasValue) {
                        SetCondense(font, condense.Value);
                    }

                    if (extend.HasValue) {
                        SetExtend(font, extend.Value);
                    }

                    SetVerticalTextAlignment(font, verticalTextAlignment);

                    if (fontFamily.HasValue) {
                        SetFontFamily(font, fontFamily.Value);
                    }

                    if (fontCharacterSet.HasValue) {
                        SetFontCharacterSet(font, fontCharacterSet.Value);
                    }
                });

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.FontId = fontId;
                    format.ApplyFont = true;
                });
                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies only the supplied alignment facets while preserving existing cell style facets.
        /// Null alignment values leave the corresponding alignment facet unchanged.
        /// </summary>
        internal void FormatCellAlignment(
            int row,
            int column,
            HorizontalAlignmentValues? horizontalAlignment,
            VerticalAlignmentValues? verticalAlignment,
            bool? wrapText,
            uint? textRotation = null,
            uint? indent = null,
            bool? shrinkToFit = null,
            uint? readingOrder = null) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    var existingAlignment = format.Alignment != null
                        ? (Alignment)format.Alignment.CloneNode(true)
                        : new Alignment();
                    if (horizontalAlignment.HasValue) {
                        existingAlignment.Horizontal = horizontalAlignment.Value;
                    }

                    if (verticalAlignment.HasValue) {
                        existingAlignment.Vertical = verticalAlignment.Value;
                    }

                    if (wrapText.HasValue) {
                        existingAlignment.WrapText = wrapText.Value;
                    }

                    if (textRotation.HasValue) {
                        existingAlignment.TextRotation = textRotation.Value;
                    }

                    if (indent.HasValue) {
                        existingAlignment.Indent = indent.Value;
                    }

                    if (shrinkToFit.HasValue) {
                        existingAlignment.ShrinkToFit = shrinkToFit.Value;
                    }

                    if (readingOrder.HasValue) {
                        existingAlignment.ReadingOrder = readingOrder.Value;
                    }

                    format.Alignment = existingAlignment;
                    format.ApplyAlignment = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies a cell fill pattern while preserving existing cell style facets.
        /// Null colors leave the corresponding fill color unset.
        /// </summary>
        internal void FormatCellFill(
            int row,
            int column,
            PatternValues pattern,
            string? foregroundColor,
            string? backgroundColor) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                var patternFill = new PatternFill {
                    PatternType = pattern
                };

                if (!string.IsNullOrWhiteSpace(foregroundColor)) {
                    patternFill.ForegroundColor = new ForegroundColor {
                        Rgb = NormalizeHexColor(foregroundColor!)
                    };
                }

                if (!string.IsNullOrWhiteSpace(backgroundColor)) {
                    patternFill.BackgroundColor = new BackgroundColor {
                        Rgb = NormalizeHexColor(backgroundColor!)
                    };
                }

                uint fillId = GetOrCreateFill(stylesheet, new Fill(patternFill));
                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.FillId = fillId;
                    format.ApplyFill = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies side-specific border facets while preserving existing cell style facets.
        /// Null border styles leave the corresponding side unchanged.
        /// </summary>
        internal void FormatCellBorder(
            int row,
            int column,
            BorderStyleValues? leftStyle,
            string? leftColor,
            BorderStyleValues? rightStyle,
            string? rightColor,
            BorderStyleValues? topStyle,
            string? topColor,
            BorderStyleValues? bottomStyle,
            string? bottomColor,
            BorderStyleValues? diagonalStyle,
            string? diagonalColor,
            bool diagonalUp,
            bool diagonalDown) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                var baseFormat = GetBaseCellFormat(stylesheet, cell.StyleIndex?.Value ?? 0U);
                uint borderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(baseFormat.BorderId), border => {
                    if (leftStyle.HasValue) border.LeftBorder = CreateBorderSide<LeftBorder>(leftStyle.Value, NormalizeOptionalHexColor(leftColor));
                    if (rightStyle.HasValue) border.RightBorder = CreateBorderSide<RightBorder>(rightStyle.Value, NormalizeOptionalHexColor(rightColor));
                    if (topStyle.HasValue) border.TopBorder = CreateBorderSide<TopBorder>(topStyle.Value, NormalizeOptionalHexColor(topColor));
                    if (bottomStyle.HasValue) border.BottomBorder = CreateBorderSide<BottomBorder>(bottomStyle.Value, NormalizeOptionalHexColor(bottomColor));
                    if (diagonalStyle.HasValue && (diagonalUp || diagonalDown)) {
                        border.DiagonalBorder = CreateBorderSide<DiagonalBorder>(diagonalStyle.Value, NormalizeOptionalHexColor(diagonalColor));
                        border.DiagonalUp = diagonalUp;
                        border.DiagonalDown = diagonalDown;
                    }
                });
                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.BorderId = borderId;
                    format.ApplyBorder = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies cell protection facets while preserving existing cell style facets.
        /// </summary>
        internal void FormatCellProtection(int row, int column, bool locked, bool formulaHidden) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.Protection = new Protection {
                        Locked = locked,
                        Hidden = formulaHidden
                    };
                    format.ApplyProtection = true;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        /// <summary>
        /// Applies the Open XML quote-prefix style flag while preserving existing cell style facets.
        /// </summary>
        internal void FormatCellQuotePrefix(int row, int column, bool quotePrefix) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                var workbookPart = _excelDocument.WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is null");
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                var stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
                EnsureDefaultStylePrimitives(stylesheet);

                ApplyCellFormatOverride(stylesheet, cell, format => {
                    format.QuotePrefix = quotePrefix;
                });

                stylesPart.Stylesheet.Save();
            });
        }

        private static void ApplyLegacyNumberFormat(Stylesheet stylesheet, CellFormat candidate, LegacyXlsCellFormat format) {
            if (format.NumberFormatCode == null || format.NumberFormatId == 0) {
                return;
            }

            candidate.NumberFormatId = format.IsBuiltInNumberFormat
                ? format.NumberFormatId
                : GetOrCreateNumberFormatId(stylesheet, format.NumberFormatCode);
            candidate.ApplyNumberFormat = true;
        }

        private static void ApplyLegacyFont(Stylesheet stylesheet, CellFormat candidate, LegacyXlsWorkbook workbook, LegacyXlsCellFormat format) {
            LegacyXlsFont? font = workbook.GetFont(format.FontIndex);
            if (font == null) {
                return;
            }

            workbook.TryResolveColor(font.ColorIndex, out string? fontColor);
            VerticalAlignmentRunValues? verticalTextAlignment = ToLegacyVerticalTextAlignment(font.Escapement);
            UnderlineValues? underlineStyle = ToLegacyUnderlineStyle(font.UnderlineStyle);
            bool hasFontMetadata = font.Family != 0 || font.CharacterSet != 1 || font.Outline || font.Shadow || font.Condense || font.Extend;
            if (font.Name == null && !font.Size.HasValue && fontColor == null && !font.Bold && !font.Italic && !font.Underline && !font.Strikeout && !verticalTextAlignment.HasValue && !hasFontMetadata) {
                return;
            }

            candidate.FontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(candidate.FontId), projectedFont => {
                if (!string.IsNullOrWhiteSpace(font.Name)) {
                    SetFontName(projectedFont, font.Name!);
                }

                if (font.Size.HasValue && font.Size.Value > 0) {
                    SetFontSize(projectedFont, font.Size.Value);
                }

                if (!string.IsNullOrWhiteSpace(fontColor)) {
                    SetFontColor(projectedFont, fontColor!);
                }

                SetBold(projectedFont, font.Bold);
                SetItalic(projectedFont, font.Italic);
                SetUnderline(projectedFont, underlineStyle);
                SetStrike(projectedFont, font.Strikeout);
                SetOutline(projectedFont, font.Outline);
                SetShadow(projectedFont, font.Shadow);
                SetCondense(projectedFont, font.Condense);
                SetExtend(projectedFont, font.Extend);
                SetVerticalTextAlignment(projectedFont, verticalTextAlignment);
                SetFontFamily(projectedFont, font.Family);
                SetFontCharacterSet(projectedFont, font.CharacterSet);
            });
            candidate.ApplyFont = true;
        }

        private static void ApplyLegacyFill(Stylesheet stylesheet, CellFormat candidate, LegacyXlsWorkbook workbook, LegacyXlsCellFormat format) {
            if (format.FillPattern == 0) {
                return;
            }

            PatternValues? pattern = ToLegacyFillPattern(format.FillPattern);
            if (!pattern.HasValue) {
                return;
            }

            string? foregroundColor = ResolveLegacyColor(workbook, format.FillForegroundColorIndex);
            string? backgroundColor = ResolveLegacyColor(workbook, format.FillBackgroundColorIndex);
            if (foregroundColor == null && backgroundColor == null) {
                return;
            }

            var patternFill = new PatternFill {
                PatternType = pattern.Value
            };

            if (!string.IsNullOrWhiteSpace(foregroundColor)) {
                patternFill.ForegroundColor = new ForegroundColor {
                    Rgb = NormalizeHexColor(foregroundColor!)
                };
            }

            if (!string.IsNullOrWhiteSpace(backgroundColor)) {
                patternFill.BackgroundColor = new BackgroundColor {
                    Rgb = NormalizeHexColor(backgroundColor!)
                };
            }

            if (format.FillPattern == 1 && foregroundColor != null) {
                patternFill.BackgroundColor = new BackgroundColor {
                    Rgb = NormalizeHexColor(foregroundColor)
                };
            }

            candidate.FillId = GetOrCreateFill(stylesheet, new Fill(patternFill));
            candidate.ApplyFill = true;
        }

        private static void ApplyLegacyAlignment(CellFormat candidate, LegacyXlsCellFormat format) {
            if (!format.ApplyAlignment) {
                return;
            }

            var alignment = candidate.Alignment != null
                ? (Alignment)candidate.Alignment.CloneNode(true)
                : new Alignment();

            HorizontalAlignmentValues? horizontal = ToLegacyHorizontalAlignment(format.HorizontalAlignment);
            if (horizontal.HasValue) {
                alignment.Horizontal = horizontal.Value;
            }

            VerticalAlignmentValues? vertical = ToLegacyVerticalAlignment(format.VerticalAlignment);
            if (vertical.HasValue) {
                alignment.Vertical = vertical.Value;
            }

            alignment.WrapText = format.WrapText;
            uint? textRotation = ToLegacyTextRotation(format.TextRotation);
            if (textRotation.HasValue) {
                alignment.TextRotation = textRotation.Value;
            }

            alignment.Indent = format.Indent;
            alignment.ShrinkToFit = format.ShrinkToFit;
            uint? readingOrder = ToLegacyReadingOrder(format.ReadingOrder);
            if (readingOrder.HasValue) {
                alignment.ReadingOrder = readingOrder.Value;
            }

            candidate.Alignment = alignment;
            candidate.ApplyAlignment = true;
        }

        private static void ApplyLegacyBorder(Stylesheet stylesheet, CellFormat candidate, LegacyXlsWorkbook workbook, LegacyXlsCellFormat format) {
            if (format.Border == null) {
                return;
            }

            LegacyXlsBorder legacyBorder = format.Border;
            candidate.BorderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(candidate.BorderId), border => {
                BorderStyleValues? leftStyle = ToLegacyBorderStyle(legacyBorder.LeftStyle);
                if (leftStyle.HasValue) border.LeftBorder = CreateBorderSide<LeftBorder>(leftStyle.Value, NormalizeOptionalHexColor(ResolveLegacyColor(workbook, legacyBorder.LeftColorIndex)));

                BorderStyleValues? rightStyle = ToLegacyBorderStyle(legacyBorder.RightStyle);
                if (rightStyle.HasValue) border.RightBorder = CreateBorderSide<RightBorder>(rightStyle.Value, NormalizeOptionalHexColor(ResolveLegacyColor(workbook, legacyBorder.RightColorIndex)));

                BorderStyleValues? topStyle = ToLegacyBorderStyle(legacyBorder.TopStyle);
                if (topStyle.HasValue) border.TopBorder = CreateBorderSide<TopBorder>(topStyle.Value, NormalizeOptionalHexColor(ResolveLegacyColor(workbook, legacyBorder.TopColorIndex)));

                BorderStyleValues? bottomStyle = ToLegacyBorderStyle(legacyBorder.BottomStyle);
                if (bottomStyle.HasValue) border.BottomBorder = CreateBorderSide<BottomBorder>(bottomStyle.Value, NormalizeOptionalHexColor(ResolveLegacyColor(workbook, legacyBorder.BottomColorIndex)));

                BorderStyleValues? diagonalStyle = ToLegacyBorderStyle(legacyBorder.DiagonalStyle);
                if (diagonalStyle.HasValue && (legacyBorder.DiagonalUp || legacyBorder.DiagonalDown)) {
                    border.DiagonalBorder = CreateBorderSide<DiagonalBorder>(diagonalStyle.Value, NormalizeOptionalHexColor(ResolveLegacyColor(workbook, legacyBorder.DiagonalColorIndex)));
                    border.DiagonalUp = legacyBorder.DiagonalUp;
                    border.DiagonalDown = legacyBorder.DiagonalDown;
                }
            });
            candidate.ApplyBorder = true;
        }

        private static void ApplyLegacyProtection(CellFormat candidate, LegacyXlsCellFormat format) {
            if (!format.ApplyProtection) {
                return;
            }

            candidate.Protection = new Protection {
                Locked = format.Locked,
                Hidden = format.FormulaHidden
            };
            candidate.ApplyProtection = true;
        }

        private static void ApplyLegacyQuotePrefix(CellFormat candidate, LegacyXlsCellFormat format) {
            if (format.QuotePrefix) {
                candidate.QuotePrefix = true;
            }
        }

        private static void ApplyLegacyCellStyleExtensions(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellFormat format) {
            foreach (LegacyXlsCellStyleExtension extension in workbook.CellStyleExtensions.Where(extension =>
                extension.HasProjectableFormatting && AppliesToEffectiveLegacyFormat(extension, format))) {
                foreach (LegacyXlsCellStyleExtensionProperty property in extension.Properties) {
                    ApplyLegacyCellStyleExtensionProperty(stylesheet, candidate, workbook, property);
                }
            }
        }

        private static bool AppliesToEffectiveLegacyFormat(LegacyXlsCellStyleExtension extension, LegacyXlsCellFormat format) {
            if (extension.AppliesToFormatIndex(format.StyleIndex)) {
                return true;
            }

            return !format.IsStyle
                && format.ParentStyleIndex != format.StyleIndex
                && extension.AppliesToFormatIndex(format.ParentStyleIndex);
        }

        private static void ApplyLegacyCellStyleExtensionProperty(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property) {
            if (property.UsesStyleXfPropMapping) {
                ApplyLegacyStyleXfProperty(stylesheet, candidate, workbook, property);
                return;
            }

            switch (property.PropertyType) {
                case 0x0004:
                    ApplyLegacyFillExtensionColor<ForegroundColor>(stylesheet, candidate, workbook, property, foreground: true);
                    break;
                case 0x0005:
                    ApplyLegacyFillExtensionColor<BackgroundColor>(stylesheet, candidate, workbook, property, foreground: false);
                    break;
                case 0x0007:
                    ApplyLegacyBorderExtensionColor(stylesheet, candidate, workbook, property, BorderColorTarget.Top);
                    break;
                case 0x0008:
                    ApplyLegacyBorderExtensionColor(stylesheet, candidate, workbook, property, BorderColorTarget.Bottom);
                    break;
                case 0x0009:
                    ApplyLegacyBorderExtensionColor(stylesheet, candidate, workbook, property, BorderColorTarget.Left);
                    break;
                case 0x000A:
                    ApplyLegacyBorderExtensionColor(stylesheet, candidate, workbook, property, BorderColorTarget.Right);
                    break;
                case 0x000B:
                    ApplyLegacyBorderExtensionColor(stylesheet, candidate, workbook, property, BorderColorTarget.Diagonal);
                    break;
                case 0x000D:
                    ApplyLegacyFontExtensionColor(stylesheet, candidate, workbook, property);
                    break;
                case 0x000E:
                    ApplyLegacyFontScheme(stylesheet, candidate, property);
                    break;
                case 0x000F:
                    ApplyLegacyExtensionIndent(candidate, property);
                    break;
            }
        }

        private static void ApplyLegacyStyleXfProperty(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property) {
            switch (property.PropertyType) {
                case 0x0000:
                    ApplyLegacyFillExtensionPattern(stylesheet, candidate, property);
                    break;
                case 0x0001:
                    ApplyLegacyFillExtensionColor<ForegroundColor>(stylesheet, candidate, workbook, property, foreground: true);
                    break;
                case 0x0002:
                    ApplyLegacyFillExtensionColor<BackgroundColor>(stylesheet, candidate, workbook, property, foreground: false);
                    break;
                case 0x0005:
                    ApplyLegacyFontExtensionColor(stylesheet, candidate, workbook, property);
                    break;
                case 0x0006:
                    ApplyLegacyBorderExtension(stylesheet, candidate, workbook, property, BorderColorTarget.Top);
                    break;
                case 0x0007:
                    ApplyLegacyBorderExtension(stylesheet, candidate, workbook, property, BorderColorTarget.Bottom);
                    break;
                case 0x0008:
                    ApplyLegacyBorderExtension(stylesheet, candidate, workbook, property, BorderColorTarget.Left);
                    break;
                case 0x0009:
                    ApplyLegacyBorderExtension(stylesheet, candidate, workbook, property, BorderColorTarget.Right);
                    break;
                case 0x000A:
                    ApplyLegacyBorderExtension(stylesheet, candidate, workbook, property, BorderColorTarget.Diagonal);
                    break;
                case 0x0012:
                    ApplyLegacyExtensionIndent(candidate, property);
                    break;
                case 0x0025:
                    ApplyLegacyFontScheme(stylesheet, candidate, property);
                    break;
            }
        }

        private static void ApplyLegacyFontExtensionColor(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property) {
            if (!TryCreateLegacyExtensionColor<DocumentFormat.OpenXml.Spreadsheet.Color>(workbook, property, out var color) || color == null) {
                return;
            }

            candidate.FontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(candidate.FontId), font => {
                SetFontColorElement(font, color);
            });
            candidate.ApplyFont = true;
        }

        private static void ApplyLegacyFontScheme(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsCellStyleExtensionProperty property) {
            FontSchemeValues? scheme = property.NumericValue switch {
                0x0001 => FontSchemeValues.Major,
                0x0002 => FontSchemeValues.Minor,
                _ => null
            };

            candidate.FontId = GetOrCreateFontVariant(stylesheet, GetOptionalValue(candidate.FontId), font => {
                SetFontSchemeElement(font, scheme);
            });
            candidate.ApplyFont = true;
        }

        private static void ApplyLegacyFillExtensionPattern(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsCellStyleExtensionProperty property) {
            if (!property.NumericValue.HasValue) {
                return;
            }

            PatternValues? pattern = ToLegacyFillPattern(checked((byte)property.NumericValue.Value));
            if (!pattern.HasValue) {
                return;
            }

            candidate.FillId = GetOrCreateFillVariant(stylesheet, GetOptionalValue(candidate.FillId), fill => {
                PatternFill patternFill = fill.PatternFill ??= new PatternFill();
                patternFill.PatternType = pattern.Value;
            });
            candidate.ApplyFill = true;
        }

        private static void ApplyLegacyFillExtensionColor<TColor>(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property,
            bool foreground) where TColor : ColorType, new() {
            if (!TryCreateLegacyExtensionColor<TColor>(workbook, property, out var color) || color == null) {
                return;
            }

            candidate.FillId = GetOrCreateFillVariant(stylesheet, GetOptionalValue(candidate.FillId), fill => {
                PatternFill patternFill = fill.PatternFill ??= new PatternFill {
                    PatternType = PatternValues.Solid
                };

                if (foreground) {
                    patternFill.ForegroundColor = color as ForegroundColor;
                } else {
                    patternFill.BackgroundColor = color as BackgroundColor;
                }
            });
            candidate.ApplyFill = true;
        }

        private static void ApplyLegacyBorderExtensionColor(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property,
            BorderColorTarget target) {
            if (!TryCreateLegacyExtensionColor<DocumentFormat.OpenXml.Spreadsheet.Color>(workbook, property, out var color) || color == null) {
                return;
            }

            candidate.BorderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(candidate.BorderId), border => {
                switch (target) {
                    case BorderColorTarget.Top:
                        SetBorderColor(border.TopBorder, color);
                        break;
                    case BorderColorTarget.Bottom:
                        SetBorderColor(border.BottomBorder, color);
                        break;
                    case BorderColorTarget.Left:
                        SetBorderColor(border.LeftBorder, color);
                        break;
                    case BorderColorTarget.Right:
                        SetBorderColor(border.RightBorder, color);
                        break;
                    case BorderColorTarget.Diagonal:
                        SetBorderColor(border.DiagonalBorder, color);
                        break;
                }
            });
            candidate.ApplyBorder = true;
        }

        private static void ApplyLegacyBorderExtension(
            Stylesheet stylesheet,
            CellFormat candidate,
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property,
            BorderColorTarget target) {
            if (!property.BorderStyle.HasValue
                || !TryCreateLegacyExtensionColor<DocumentFormat.OpenXml.Spreadsheet.Color>(workbook, property, out var color)
                || color == null) {
                return;
            }

            BorderStyleValues? style = ToLegacyBorderStyle(checked((byte)property.BorderStyle.Value));
            if (!style.HasValue) {
                return;
            }

            candidate.BorderId = GetOrCreateBorderVariant(stylesheet, GetOptionalValue(candidate.BorderId), border => {
                switch (target) {
                    case BorderColorTarget.Top:
                        SetBorder(border.TopBorder ??= new TopBorder(), style.Value, color);
                        break;
                    case BorderColorTarget.Bottom:
                        SetBorder(border.BottomBorder ??= new BottomBorder(), style.Value, color);
                        break;
                    case BorderColorTarget.Left:
                        SetBorder(border.LeftBorder ??= new LeftBorder(), style.Value, color);
                        break;
                    case BorderColorTarget.Right:
                        SetBorder(border.RightBorder ??= new RightBorder(), style.Value, color);
                        break;
                    case BorderColorTarget.Diagonal:
                        SetBorder(border.DiagonalBorder ??= new DiagonalBorder(), style.Value, color);
                        break;
                }
            });
            candidate.ApplyBorder = true;
        }

        private static void ApplyLegacyExtensionIndent(CellFormat candidate, LegacyXlsCellStyleExtensionProperty property) {
            if (!property.NumericValue.HasValue) {
                return;
            }

            var alignment = candidate.Alignment != null
                ? (Alignment)candidate.Alignment.CloneNode(true)
                : new Alignment();
            alignment.Indent = property.NumericValue.Value;
            candidate.Alignment = alignment;
            candidate.ApplyAlignment = true;
        }

        private static bool TryCreateLegacyExtensionColor<TColor>(
            LegacyXlsWorkbook workbook,
            LegacyXlsCellStyleExtensionProperty property,
            out TColor? color) where TColor : ColorType, new() {
            color = null;
            if (!property.ColorType.HasValue
                || !property.ColorValue.HasValue) {
                return false;
            }

            color = new TColor();
            switch (property.ColorType.Value) {
                case 0x0001:
                    string? indexedColor = ResolveLegacyColor(workbook, checked((ushort)property.ColorValue.Value));
                    if (string.IsNullOrWhiteSpace(indexedColor)) {
                        color = null;
                        return false;
                    }

                    color.Rgb = NormalizeHexColor(indexedColor!);
                    break;
                case 0x0002:
                    color.Rgb = NormalizeHexColor(ToLegacyExtensionArgb(property.ColorValue.Value));
                    break;
                case 0x0003:
                    if (!TryMapLegacyThemeColor(property.ColorValue.Value, out uint openXmlTheme)) {
                        color = null;
                        return false;
                    }

                    color.Theme = openXmlTheme;
                    break;
                default:
                    color = null;
                    return false;
            }

            double tint = ToLegacyExtensionTint(property.ColorTintShade.GetValueOrDefault());
            if (Math.Abs(tint) > double.Epsilon) {
                color.Tint = tint;
            }

            return true;
        }

        private static uint GetOrCreateFillVariant(Stylesheet stylesheet, uint? baseFillId, Action<Fill> mutate) {
            var fills = stylesheet.Fills ??= new Fills();
            var baseFill = fills.Elements<Fill>().ElementAtOrDefault((int)(baseFillId ?? 0U));
            var candidate = baseFill != null
                ? (Fill)baseFill.CloneNode(true)
                : new Fill();

            mutate(candidate);
            return GetOrCreateFill(stylesheet, candidate);
        }

        private static void SetBorderColor(BorderPropertiesType? borderSide, DocumentFormat.OpenXml.Spreadsheet.Color color) {
            if (borderSide?.Style == null) {
                return;
            }

            foreach (DocumentFormat.OpenXml.Spreadsheet.Color existing in borderSide.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList()) {
                existing.Remove();
            }

            borderSide.Append((DocumentFormat.OpenXml.Spreadsheet.Color)color.CloneNode(true));
        }

        private static void SetBorder(BorderPropertiesType borderSide, BorderStyleValues style, DocumentFormat.OpenXml.Spreadsheet.Color color) {
            borderSide.Style = style;
            SetBorderColor(borderSide, color);
        }

        private static string ToLegacyExtensionArgb(uint longRgba) {
            byte red = (byte)(longRgba & 0xff);
            byte green = (byte)((longRgba >> 8) & 0xff);
            byte blue = (byte)((longRgba >> 16) & 0xff);
            byte alpha = (byte)((longRgba >> 24) & 0xff);
            return string.Concat(
                alpha.ToString("X2", CultureInfo.InvariantCulture),
                red.ToString("X2", CultureInfo.InvariantCulture),
                green.ToString("X2", CultureInfo.InvariantCulture),
                blue.ToString("X2", CultureInfo.InvariantCulture));
        }

        private static bool TryMapLegacyThemeColor(uint legacyThemeColor, out uint openXmlThemeColor) {
            switch (legacyThemeColor) {
                case 0:
                    openXmlThemeColor = 1;
                    return true;
                case 1:
                    openXmlThemeColor = 0;
                    return true;
                case 2:
                    openXmlThemeColor = 3;
                    return true;
                case 3:
                    openXmlThemeColor = 2;
                    return true;
                case >= 4 and <= 11:
                    openXmlThemeColor = legacyThemeColor;
                    return true;
                default:
                    openXmlThemeColor = 0;
                    return false;
            }
        }

        private static double ToLegacyExtensionTint(short tintShade) {
            return tintShade == 0 ? 0D : Math.Round(tintShade / 32767D, 5, MidpointRounding.AwayFromZero);
        }

        private enum BorderColorTarget {
            Top,
            Bottom,
            Left,
            Right,
            Diagonal
        }

        private static HorizontalAlignmentValues? ToLegacyHorizontalAlignment(byte alignment) {
            return alignment switch {
                1 => HorizontalAlignmentValues.Left,
                2 => HorizontalAlignmentValues.Center,
                3 => HorizontalAlignmentValues.Right,
                4 => HorizontalAlignmentValues.Fill,
                5 => HorizontalAlignmentValues.Justify,
                6 => HorizontalAlignmentValues.CenterContinuous,
                7 => HorizontalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static VerticalAlignmentValues? ToLegacyVerticalAlignment(byte alignment) {
            return alignment switch {
                0 => VerticalAlignmentValues.Top,
                1 => VerticalAlignmentValues.Center,
                2 => VerticalAlignmentValues.Bottom,
                3 => VerticalAlignmentValues.Justify,
                4 => VerticalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static uint? ToLegacyTextRotation(byte rotation) {
            return rotation <= 180 || rotation == 255 ? rotation : null;
        }

        private static uint? ToLegacyReadingOrder(byte readingOrder) {
            return readingOrder <= 2 ? readingOrder : null;
        }

        private static BorderStyleValues? ToLegacyBorderStyle(byte style) {
            return style switch {
                1 => BorderStyleValues.Thin,
                2 => BorderStyleValues.Medium,
                3 => BorderStyleValues.Dashed,
                4 => BorderStyleValues.Dotted,
                5 => BorderStyleValues.Thick,
                6 => BorderStyleValues.Double,
                7 => BorderStyleValues.Hair,
                8 => BorderStyleValues.MediumDashed,
                9 => BorderStyleValues.DashDot,
                10 => BorderStyleValues.MediumDashDot,
                11 => BorderStyleValues.DashDotDot,
                12 => BorderStyleValues.MediumDashDotDot,
                13 => BorderStyleValues.SlantDashDot,
                _ => null
            };
        }

        private static VerticalAlignmentRunValues? ToLegacyVerticalTextAlignment(LegacyXlsFontEscapement escapement) {
            return escapement == LegacyXlsFontEscapement.Superscript
                ? VerticalAlignmentRunValues.Superscript
                : escapement == LegacyXlsFontEscapement.Subscript
                    ? VerticalAlignmentRunValues.Subscript
                    : null;
        }

        private static UnderlineValues? ToLegacyUnderlineStyle(byte underlineStyle) {
            return underlineStyle switch {
                0x01 => UnderlineValues.Single,
                0x02 => UnderlineValues.Double,
                0x21 => UnderlineValues.SingleAccounting,
                0x22 => UnderlineValues.DoubleAccounting,
                _ => null
            };
        }

        private static PatternValues? ToLegacyFillPattern(byte pattern) {
            return pattern switch {
                1 => PatternValues.Solid,
                2 => PatternValues.MediumGray,
                3 => PatternValues.DarkGray,
                4 => PatternValues.LightGray,
                5 => PatternValues.DarkHorizontal,
                6 => PatternValues.DarkVertical,
                7 => PatternValues.DarkDown,
                8 => PatternValues.DarkUp,
                9 => PatternValues.DarkGrid,
                10 => PatternValues.DarkTrellis,
                11 => PatternValues.LightHorizontal,
                12 => PatternValues.LightVertical,
                13 => PatternValues.LightDown,
                14 => PatternValues.LightUp,
                15 => PatternValues.LightGrid,
                16 => PatternValues.LightTrellis,
                17 => PatternValues.Gray125,
                18 => PatternValues.Gray0625,
                _ => null
            };
        }

        private static string? ResolveLegacyColor(LegacyXlsWorkbook workbook, ushort colorIndex) {
            return workbook.TryResolveColor(colorIndex, out string? color) ? color : null;
        }
    }
}
