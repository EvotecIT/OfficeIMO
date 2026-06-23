using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

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

                CellFormat candidate = GetBaseCellFormat(stylesheet, 0U);
                ApplyLegacyNumberFormat(stylesheet, candidate, format);
                ApplyLegacyFont(stylesheet, candidate, workbook, format);
                ApplyLegacyFill(stylesheet, candidate, workbook, format);
                ApplyLegacyAlignment(candidate, format);
                ApplyLegacyBorder(stylesheet, candidate, workbook, format);
                ApplyLegacyProtection(candidate, format);
                ApplyLegacyQuotePrefix(candidate, format);

                styleIndex = AppendOrReuseCellFormat(stylesheet, candidate);
                stylesPart.Stylesheet.Save();
            });

            return styleIndex;
        }

        /// <summary>
        /// Writes an imported legacy XLS error value as an Open XML error cell.
        /// </summary>
        internal void SetLegacyErrorCellValue(int row, int column, string errorText) {
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
            VerticalAlignmentRunValues? verticalTextAlignment = null) {
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

                    if (underline.HasValue) {
                        SetUnderline(font, underline.Value);
                    }

                    if (strike.HasValue) {
                        SetStrike(font, strike.Value);
                    }

                    SetVerticalTextAlignment(font, verticalTextAlignment);
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
            if (font.Name == null && !font.Size.HasValue && fontColor == null && !font.Bold && !font.Italic && !font.Underline && !font.Strikeout) {
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
                SetUnderline(projectedFont, font.Underline);
                SetStrike(projectedFont, font.Strikeout);
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
