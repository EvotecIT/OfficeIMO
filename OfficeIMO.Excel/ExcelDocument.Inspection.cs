using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Creates an immutable workbook snapshot that exposes OfficeIMO's current workbook model for downstream integrations.
        /// </summary>
        public ExcelWorkbookSnapshot CreateInspectionSnapshot(ExcelReadOptions? options = null) {
            var effectiveOptions = options ?? new ExcelReadOptions {
                UseCachedFormulaResult = true,
                TreatDatesUsingNumberFormat = true,
            };

            var snapshot = new ExcelWorkbookSnapshot {
                FilePath = string.IsNullOrWhiteSpace(FilePath) ? null : FilePath,
            };

            try {
                snapshot.Title = _spreadSheetDocument.PackageProperties.Title;
            } catch {
                snapshot.Title = null;
            }

            using (var reader = CreateReader(effectiveOptions)) {
                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
                var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
                var styleContext = StyleInspectionContext.Create(workbookPart.WorkbookStylesPart?.Stylesheet);
                var sheetElements = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();

                for (int sheetIndex = 0; sheetIndex < sheetElements.Count; sheetIndex++) {
                    var sheet = sheetElements[sheetIndex];
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                    var sheetName = sheet.Name?.Value ?? $"Sheet{sheetIndex + 1}";
                    var readerSheet = reader.GetSheet(sheetName);
                    var typedValues = BuildTypedCellMap(readerSheet);
                    var hyperlinkMap = BuildHyperlinkMap(worksheetPart);
                    var commentMap = BuildCommentMap(worksheetPart);

                    var worksheetSnapshot = new ExcelWorksheetSnapshot {
                        Name = sheetName,
                        Index = sheetIndex,
                        Hidden = sheet.State?.Value == SheetStateValues.Hidden || sheet.State?.Value == SheetStateValues.VeryHidden,
                        RightToLeft = worksheetPart.Worksheet
                            .GetFirstChild<SheetViews>()?
                            .Elements<SheetView>()
                            .FirstOrDefault()?
                            .RightToLeft?.Value == true,
                        TabColorArgb = GetColorArgb(worksheetPart.Worksheet.GetFirstChild<SheetProperties>()?.TabColor),
                        UsedRangeA1 = readerSheet.GetUsedRangeA1(),
                    };

                    var pane = worksheetPart.Worksheet
                        .GetFirstChild<SheetViews>()?
                        .Elements<SheetView>()
                        .FirstOrDefault()?
                        .GetFirstChild<Pane>();

                    if (pane != null && (pane.State?.Value == PaneStateValues.Frozen || pane.State?.Value == PaneStateValues.FrozenSplit)) {
                        worksheetSnapshot.FrozenRowCount = ConvertToInt(pane.VerticalSplit);
                        worksheetSnapshot.FrozenColumnCount = ConvertToInt(pane.HorizontalSplit);
                    }

                    var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                    if (columns != null) {
                        foreach (var column in columns.Elements<Column>()) {
                            var min = checked((int)(column.Min?.Value ?? 0U));
                            var max = checked((int)(column.Max?.Value ?? 0U));
                            if (min <= 0 || max <= 0 || max < min) {
                                continue;
                            }

                            worksheetSnapshot.AddColumn(new ExcelColumnSnapshot {
                                StartIndex = min,
                                EndIndex = max,
                                Width = column.Width?.Value,
                                Hidden = column.Hidden?.Value == true,
                                CustomWidth = column.CustomWidth?.Value == true,
                            });
                        }
                    }

                    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        foreach (var row in sheetData.Elements<Row>()) {
                            var rowIndex = checked((int)(row.RowIndex?.Value ?? 0U));
                            if (rowIndex > 0 && (row.Hidden?.Value == true || row.CustomHeight?.Value == true || row.Height != null)) {
                                worksheetSnapshot.AddRow(new ExcelRowSnapshot {
                                    Index = rowIndex,
                                    Height = row.Height?.Value,
                                    Hidden = row.Hidden?.Value == true,
                                    CustomHeight = row.CustomHeight?.Value == true,
                                });
                            }

                            foreach (var cell in row.Elements<Cell>()) {
                                if (string.IsNullOrWhiteSpace(cell.CellReference?.Value)) {
                                    continue;
                                }

                                var cellReference = cell.CellReference!.Value!;
                                var (_, columnIndex) = A1.ParseCellRef(cellReference);
                                if (rowIndex <= 0 || columnIndex <= 0) {
                                    continue;
                                }

                                worksheetSnapshot.AddCell(new ExcelCellSnapshot {
                                    Row = rowIndex,
                                    Column = columnIndex,
                                    Value = GetTypedValue(typedValues, rowIndex, columnIndex),
                                    Formula = cell.CellFormula?.Text,
                                    StyleIndex = cell.StyleIndex?.Value,
                                    Style = BuildCellStyleSnapshot(styleContext, cell.StyleIndex?.Value),
                                    Hyperlink = hyperlinkMap.TryGetValue(cellReference, out var hyperlink) ? hyperlink : null,
                                    Comment = commentMap.TryGetValue(cellReference, out var comment) ? comment : null,
                                });
                            }
                        }
                    }

                    var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
                    if (mergeCells != null) {
                        foreach (var mergeCell in mergeCells.Elements<MergeCell>()) {
                            var reference = mergeCell.Reference?.Value;
                            if (string.IsNullOrWhiteSpace(reference)) {
                                continue;
                            }

                            int rowStart;
                            int columnStart;
                            int rowEnd;
                            int columnEnd;
                            if (!A1.TryParseRange(reference!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                                continue;
                            }

                            worksheetSnapshot.AddMergedRange(new ExcelMergedRangeSnapshot {
                                A1Range = reference!,
                                StartRow = rowStart,
                                EndRow = rowEnd,
                                StartColumn = columnStart,
                                EndColumn = columnEnd,
                            });
                        }
                    }

                    worksheetSnapshot.AutoFilter = BuildAutoFilterSnapshot(
                        worksheetPart.Worksheet.Elements<AutoFilter>().FirstOrDefault());
                    worksheetSnapshot.Protection = BuildWorksheetProtectionSnapshot(
                        worksheetPart.Worksheet.Elements<SheetProtection>().FirstOrDefault());
                    foreach (var validation in BuildDataValidationSnapshots(worksheetPart.Worksheet.GetFirstChild<DataValidations>())) {
                        worksheetSnapshot.AddValidation(validation);
                    }

                    foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                        var table = tableDefinitionPart.Table;
                        var tableSnapshot = BuildTableSnapshot(table);
                        if (tableSnapshot != null) {
                            worksheetSnapshot.AddTable(tableSnapshot);
                        }
                    }

                    snapshot.AddWorksheet(worksheetSnapshot);
                }

                var definedNames = workbook.DefinedNames;
                if (definedNames != null) {
                    foreach (var definedName in definedNames.Elements<DefinedName>()) {
                        var name = definedName.Name?.Value;
                        if (string.IsNullOrWhiteSpace(name)) {
                            continue;
                        }

                        snapshot.AddNamedRange(new ExcelNamedRangeSnapshot {
                            Name = name!,
                            ReferenceA1 = definedName.Text ?? string.Empty,
                            SheetName = ResolveLocalSheetName(sheetElements, definedName),
                            IsBuiltIn = name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase),
                        });
                    }
                }
            }

            return snapshot;
        }

        private static ExcelWorksheetProtectionSnapshot? BuildWorksheetProtectionSnapshot(SheetProtection? protection) {
            if (protection == null) {
                return null;
            }

            return new ExcelWorksheetProtectionSnapshot {
                AllowSelectLockedCells = protection.SelectLockedCells?.Value == true,
                AllowSelectUnlockedCells = protection.SelectUnlockedCells?.Value == true,
                AllowFormatCells = protection.FormatCells?.Value == true,
                AllowFormatColumns = protection.FormatColumns?.Value == true,
                AllowFormatRows = protection.FormatRows?.Value == true,
                AllowInsertColumns = protection.InsertColumns?.Value == true,
                AllowInsertRows = protection.InsertRows?.Value == true,
                AllowInsertHyperlinks = protection.InsertHyperlinks?.Value == true,
                AllowDeleteColumns = protection.DeleteColumns?.Value == true,
                AllowDeleteRows = protection.DeleteRows?.Value == true,
                AllowSort = protection.Sort?.Value == true,
                AllowAutoFilter = protection.AutoFilter?.Value == true,
                AllowPivotTables = protection.PivotTables?.Value == true,
            };
        }

        private static Dictionary<string, ExcelHyperlinkSnapshot> BuildHyperlinkMap(WorksheetPart worksheetPart) {
            var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                return new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);
            }

            var externalRelationships = worksheetPart.HyperlinkRelationships
                .Where(relationship => !string.IsNullOrWhiteSpace(relationship.Id))
                .ToDictionary(relationship => relationship.Id!, relationship => relationship, StringComparer.OrdinalIgnoreCase);
            var map = new Dictionary<string, ExcelHyperlinkSnapshot>(StringComparer.OrdinalIgnoreCase);

            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>()) {
                var reference = hyperlink.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                string? target = null;
                bool isExternal = false;

                var relationshipId = hyperlink.Id?.Value;
                if (!string.IsNullOrWhiteSpace(relationshipId) && externalRelationships.TryGetValue(relationshipId!, out var relationship)) {
                    target = relationship.Uri?.OriginalString;
                    isExternal = true;
                } else if (!string.IsNullOrWhiteSpace(hyperlink.Location?.Value)) {
                    target = hyperlink.Location!.Value!;
                }

                if (string.IsNullOrWhiteSpace(target)) {
                    continue;
                }

                map[reference!] = new ExcelHyperlinkSnapshot {
                    IsExternal = isExternal,
                    Target = target!,
                };
            }

            return map;
        }

        private static Dictionary<string, ExcelCommentSnapshot> BuildCommentMap(WorksheetPart worksheetPart) {
            var commentsPart = worksheetPart.WorksheetCommentsPart;
            var comments = commentsPart?.Comments;
            if (comments?.CommentList == null) {
                return new Dictionary<string, ExcelCommentSnapshot>(StringComparer.OrdinalIgnoreCase);
            }

            var authorNames = comments.Authors?
                .Elements<Author>()
                .Select(author => author.Text ?? string.Empty)
                .ToList()
                ?? new List<string>();

            var map = new Dictionary<string, ExcelCommentSnapshot>(StringComparer.OrdinalIgnoreCase);
            foreach (var comment in comments.CommentList.Elements<Comment>()) {
                var reference = comment.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                string? author = null;
                var authorId = comment.AuthorId?.Value;
                if (authorId.HasValue && authorId.Value < authorNames.Count) {
                    author = authorNames[checked((int)authorId.Value)];
                }

                map[reference!] = new ExcelCommentSnapshot {
                    Author = string.IsNullOrWhiteSpace(author) ? null : author,
                    Text = ExtractCommentText(comment.CommentText),
                };
            }

            return map;
        }

        private static string ExtractCommentText(CommentText? commentText) {
            if (commentText == null) {
                return string.Empty;
            }

            var builder = new System.Text.StringBuilder();
            foreach (var element in commentText.Descendants<OpenXmlElement>()) {
                if (element is Text text) {
                    builder.Append(text.Text);
                } else if (element is Break) {
                    builder.Append('\n');
                }
            }

            return builder
                .ToString()
                .Replace("\r\n", "\n")
                .Replace('\r', '\n');
        }

        private static ExcelCellStyleSnapshot? BuildCellStyleSnapshot(StyleInspectionContext context, uint? styleIndex) {
            if (!styleIndex.HasValue) {
                return null;
            }

            var cellFormat = context.GetCellFormat(styleIndex.Value);
            if (cellFormat == null) {
                return null;
            }

            var numberFormatId = cellFormat.NumberFormatId?.Value ?? 0U;
            var font = context.GetFont(cellFormat.FontId?.Value ?? 0U);
            var fill = context.GetFill(cellFormat.FillId?.Value ?? 0U);
            var border = context.GetBorder(cellFormat.BorderId?.Value ?? 0U);

            return new ExcelCellStyleSnapshot {
                StyleIndex = styleIndex.Value,
                NumberFormatId = numberFormatId,
                NumberFormatCode = context.GetNumberFormatCode(numberFormatId),
                IsDateLike = context.IsDateLike(numberFormatId),
                Bold = font?.Bold != null,
                Italic = font?.Italic != null,
                Underline = font?.Underline != null,
                FontColorArgb = GetColorArgb(font?.Color),
                FillColorArgb = GetFillColorArgb(fill),
                Border = BuildBorderSnapshot(border),
                HorizontalAlignment = cellFormat.Alignment?.Horizontal?.InnerText,
                VerticalAlignment = cellFormat.Alignment?.Vertical?.InnerText,
                WrapText = cellFormat.Alignment?.WrapText?.Value == true,
            };
        }

        private static ExcelAutoFilterSnapshot? BuildAutoFilterSnapshot(AutoFilter? autoFilter) {
            var reference = autoFilter?.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(reference!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                return null;
            }

            var snapshot = new ExcelAutoFilterSnapshot {
                A1Range = reference!,
                StartRow = rowStart,
                EndRow = rowEnd,
                StartColumn = columnStart,
                EndColumn = columnEnd,
            };

            foreach (var filterColumn in autoFilter!.Elements<FilterColumn>()) {
                var columnId = checked((int)(filterColumn.ColumnId?.Value ?? 0U));
                var columnSnapshot = new ExcelFilterColumnSnapshot {
                    ColumnId = columnId,
                };

                var filters = filterColumn.GetFirstChild<Filters>();
                if (filters != null) {
                    foreach (var filter in filters.Elements<Filter>()) {
                        var value = filter.Val?.Value;
                        if (!string.IsNullOrWhiteSpace(value)) {
                            columnSnapshot.AddValue(value!);
                        }
                    }
                }

                var customFilters = filterColumn.GetFirstChild<CustomFilters>();
                if (customFilters != null) {
                    var customFiltersSnapshot = new ExcelCustomFiltersSnapshot {
                        MatchAll = customFilters.And?.Value == true,
                    };

                    foreach (var customFilter in customFilters.Elements<CustomFilter>()) {
                        var value = customFilter.Val?.Value;
                        if (string.IsNullOrWhiteSpace(value)) {
                            continue;
                        }

                        customFiltersSnapshot.AddCondition(new ExcelCustomFilterConditionSnapshot {
                            Operator = GetOpenXmlAttributeValue(customFilter, "operator"),
                            Value = value!,
                        });
                    }

                    if (customFiltersSnapshot.Conditions.Count > 0) {
                        columnSnapshot.CustomFilters = customFiltersSnapshot;
                    }
                }

                snapshot.AddColumn(columnSnapshot);
            }

            return snapshot;
        }

        private static IReadOnlyList<ExcelDataValidationSnapshot> BuildDataValidationSnapshots(DataValidations? dataValidations) {
            var snapshots = new List<ExcelDataValidationSnapshot>();
            if (dataValidations == null) {
                return snapshots;
            }

            foreach (var dataValidation in dataValidations.Elements<DataValidation>()) {
                var snapshot = new ExcelDataValidationSnapshot {
                    Type = GetOpenXmlAttributeValue(dataValidation, "type")?.ToLowerInvariant(),
                    Operator = GetOpenXmlAttributeValue(dataValidation, "operator"),
                    AllowBlank = dataValidation.AllowBlank?.Value == true,
                    Formula1 = dataValidation.GetFirstChild<Formula1>()?.Text,
                    Formula2 = dataValidation.GetFirstChild<Formula2>()?.Text,
                };

                var sequenceOfReferences = dataValidation.SequenceOfReferences?.InnerText;
                if (!string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                    foreach (var range in sequenceOfReferences!
                                 .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(value => value.Trim())
                                 .Where(value => !string.IsNullOrWhiteSpace(value))) {
                        snapshot.AddRange(range);
                    }
                }

                if (snapshot.A1Ranges.Count > 0) {
                    snapshots.Add(snapshot);
                }
            }

            return snapshots;
        }

        private static ExcelTableSnapshot? BuildTableSnapshot(Table? table) {
            if (table == null) {
                return null;
            }

            var range = table.Reference?.Value;
            if (string.IsNullOrWhiteSpace(range)) {
                return null;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!A1.TryParseRange(range!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                return null;
            }

            var snapshot = new ExcelTableSnapshot {
                Name = table.Name?.Value ?? table.DisplayName?.Value ?? string.Empty,
                A1Range = range!,
                StartRow = rowStart,
                EndRow = rowEnd,
                StartColumn = columnStart,
                EndColumn = columnEnd,
                StyleName = table.TableStyleInfo?.Name?.Value,
                HasHeaderRow = (table.HeaderRowCount?.Value ?? 0U) > 0,
                TotalsRowShown = table.TotalsRowShown?.Value == true,
                AutoFilter = BuildAutoFilterSnapshot(table.Elements<AutoFilter>().FirstOrDefault()),
            };

            foreach (var tableColumn in table.TableColumns?.Elements<TableColumn>() ?? Enumerable.Empty<TableColumn>()) {
                snapshot.AddColumn(new ExcelTableColumnSnapshot {
                    Index = checked((int)(tableColumn.Id?.Value ?? 0U)),
                    Name = tableColumn.Name?.Value ?? string.Empty,
                    TotalsRowFunction = GetOpenXmlAttributeValue(tableColumn, "totalsRowFunction"),
                });
            }

            return snapshot;
        }

        private static string? GetOpenXmlAttributeValue(OpenXmlElement element, string localName) {
            if (element == null) throw new ArgumentNullException(nameof(element));
            if (string.IsNullOrWhiteSpace(localName)) throw new ArgumentException("Attribute name is required.", nameof(localName));

            var attribute = element.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
        }

        private static ExcelCellBorderSnapshot? BuildBorderSnapshot(Border? border) {
            if (border == null) {
                return null;
            }

            var left = BuildBorderSideSnapshot(border.LeftBorder);
            var right = BuildBorderSideSnapshot(border.RightBorder);
            var top = BuildBorderSideSnapshot(border.TopBorder);
            var bottom = BuildBorderSideSnapshot(border.BottomBorder);

            if (left == null && right == null && top == null && bottom == null) {
                return null;
            }

            return new ExcelCellBorderSnapshot {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
            };
        }

        private static ExcelBorderSideSnapshot? BuildBorderSideSnapshot(BorderPropertiesType? borderSide) {
            if (borderSide == null) {
                return null;
            }

            var style = ExtractBorderStyle(borderSide);

            var colorArgb = GetColorArgb(borderSide.GetFirstChild<Color>());

            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(colorArgb)) {
                return null;
            }

            return new ExcelBorderSideSnapshot {
                Style = style,
                ColorArgb = colorArgb,
            };
        }

        private static string? ExtractBorderStyle(BorderPropertiesType borderSide) {
            var xml = borderSide.OuterXml;
            if (string.IsNullOrWhiteSpace(xml)) {
                return null;
            }

            const string marker = "style=\"";
            var index = xml.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (index < 0) {
                return null;
            }

            index += marker.Length;
            var endIndex = xml.IndexOf('"', index);
            if (endIndex <= index) {
                return null;
            }

            var value = xml.Substring(index, endIndex - index);
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim().ToLowerInvariant();
        }

        private static string? GetFillColorArgb(Fill? fill) {
            var patternFill = fill?.PatternFill;
            if (patternFill == null) {
                return null;
            }

            return GetColorArgb(patternFill.ForegroundColor) ?? GetColorArgb(patternFill.BackgroundColor);
        }

        private static string? GetColorArgb(OpenXmlElement? colorElement) {
            string? rgb = colorElement switch {
                Color color => color.Rgb?.Value,
                TabColor tabColor => tabColor.Rgb?.Value,
                ForegroundColor foregroundColor => foregroundColor.Rgb?.Value,
                BackgroundColor backgroundColor => backgroundColor.Rgb?.Value,
                _ => null,
            };

            if (string.IsNullOrWhiteSpace(rgb)) {
                return null;
            }

            rgb = rgb!.Trim();
            if (rgb.Length == 6) {
                return "FF" + rgb.ToUpperInvariant();
            }

            return rgb.Length == 8 ? rgb.ToUpperInvariant() : null;
        }

        private static Dictionary<long, object?> BuildTypedCellMap(ExcelSheetReader readerSheet) {
            var map = new Dictionary<long, object?>();
            foreach (var cell in readerSheet.EnumerateCells()) {
                map[GetCellKey(cell.Row, cell.Column)] = cell.Value;
            }

            return map;
        }

        private static object? GetTypedValue(Dictionary<long, object?> map, int row, int column) {
            object? value;
            return map.TryGetValue(GetCellKey(row, column), out value) ? value : null;
        }

        private static long GetCellKey(int row, int column) {
            return ((long)row << 20) | (uint)column;
        }

        private static int ConvertToInt(DoubleValue? value) {
            if (value == null || !value.HasValue) {
                return 0;
            }

            return (int)Math.Round(value.Value);
        }

        private static string? ResolveLocalSheetName(IReadOnlyList<Sheet> sheetElements, DefinedName definedName) {
            if (definedName.LocalSheetId == null) {
                return null;
            }

            var index = checked((int)definedName.LocalSheetId.Value);
            if (index < 0 || index >= sheetElements.Count) {
                return null;
            }

            return sheetElements[index].Name?.Value;
        }

        private sealed class StyleInspectionContext {
            private readonly IReadOnlyList<CellFormat> _cellFormats;
            private readonly IReadOnlyList<Font> _fonts;
            private readonly IReadOnlyList<Fill> _fills;
            private readonly IReadOnlyList<Border> _borders;
            private readonly Dictionary<uint, string> _numberFormats;

            private StyleInspectionContext(
                IReadOnlyList<CellFormat> cellFormats,
                IReadOnlyList<Font> fonts,
                IReadOnlyList<Fill> fills,
                IReadOnlyList<Border> borders,
                Dictionary<uint, string> numberFormats) {
                _cellFormats = cellFormats;
                _fonts = fonts;
                _fills = fills;
                _borders = borders;
                _numberFormats = numberFormats;
            }

            internal static StyleInspectionContext Create(Stylesheet? stylesheet) {
                var numberFormats = new Dictionary<uint, string>();
                if (stylesheet?.NumberingFormats != null) {
                    foreach (var numberingFormat in stylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                        if (numberingFormat.NumberFormatId?.Value is uint id && numberingFormat.FormatCode?.Value is string code) {
                            numberFormats[id] = code;
                        }
                    }
                }

                return new StyleInspectionContext(
                    stylesheet?.CellFormats?.Elements<CellFormat>().ToList() ?? new List<CellFormat>(),
                    stylesheet?.Fonts?.Elements<Font>().ToList() ?? new List<Font>(),
                    stylesheet?.Fills?.Elements<Fill>().ToList() ?? new List<Fill>(),
                    stylesheet?.Borders?.Elements<Border>().ToList() ?? new List<Border>(),
                    numberFormats);
            }

            internal CellFormat? GetCellFormat(uint styleIndex) {
                return styleIndex < _cellFormats.Count ? _cellFormats[(int)styleIndex] : null;
            }

            internal Font? GetFont(uint fontId) {
                return fontId < _fonts.Count ? _fonts[(int)fontId] : null;
            }

            internal Fill? GetFill(uint fillId) {
                return fillId < _fills.Count ? _fills[(int)fillId] : null;
            }

            internal Border? GetBorder(uint borderId) {
                return borderId < _borders.Count ? _borders[(int)borderId] : null;
            }

            internal string? GetNumberFormatCode(uint numberFormatId) {
                string? code;
                return _numberFormats.TryGetValue(numberFormatId, out code) ? code : null;
            }

            internal bool IsDateLike(uint numberFormatId) {
                return IsBuiltInDate(numberFormatId)
                    || (_numberFormats.TryGetValue(numberFormatId, out var code) && LooksLikeDateFormat(code));
            }

            private static bool IsBuiltInDate(uint id) {
                return id is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                    or 27 or 30 or 36 or 45 or 46 or 47;
            }

            private static bool LooksLikeDateFormat(string code) {
                var lower = code.ToLowerInvariant();
                if (lower.IndexOf('d') >= 0 || lower.IndexOf('y') >= 0 || lower.IndexOf('h') >= 0 || lower.IndexOf('s') >= 0) {
                    return true;
                }

                return lower.Contains('m') && (lower.Contains('d') || lower.Contains('y') || lower.Contains('h'));
            }
        }
    }
}
