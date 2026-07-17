using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;

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
                DateSystem = DateSystem,
            };

            try {
                snapshot.Title = _spreadSheetDocument.PackageProperties.Title;
                snapshot.Author = _spreadSheetDocument.PackageProperties.Creator;
                snapshot.Subject = _spreadSheetDocument.PackageProperties.Subject;
                snapshot.Keywords = _spreadSheetDocument.PackageProperties.Keywords;
            } catch {
                snapshot.Title = null;
                snapshot.Author = null;
                snapshot.Subject = null;
                snapshot.Keywords = null;
            }

            using (var reader = CreateReader(effectiveOptions)) {
                var workbookPart = WorkbookPartRoot ?? throw new InvalidOperationException("WorkbookPart is missing.");
                var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
                var styleContext = StyleInspectionContext.Create(workbookPart, workbookPart.WorkbookStylesPart?.Stylesheet);
                var sharedStringItems = workbookPart.SharedStringTablePart?
                    .SharedStringTable?
                    .Elements<SharedStringItem>()
                    .ToList();
                snapshot.SlicerPartCount = CountPackageParts(workbookPart, IsNativeSlicerPackagePart);
                snapshot.TimelinePartCount = CountPackageParts(workbookPart, IsNativeTimelinePackagePart);
                snapshot.SlicerBindingMetadataPartCount = CountPackageParts(workbookPart, IsSlicerBindingMetadataPart);
                snapshot.TimelineBindingMetadataPartCount = CountPackageParts(workbookPart, IsTimelineBindingMetadataPart);
                snapshot.ConnectionPartCount = CountPackagePartsByContentType(workbookPart, "connections");
                snapshot.QueryTablePartCount = CountPackagePartsByContentType(workbookPart, "queryTable");
                snapshot.ChartPartCount = CountPackageParts(workbookPart, part => part is ChartPart);
                snapshot.PivotTablePartCount = CountPackageParts(workbookPart, part => part is PivotTablePart);
                var sheetElements = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                int? activeWorksheetIndex = GetActiveWorksheetIndex(workbook, sheetElements.Count);
                if (activeWorksheetIndex.HasValue) {
                    snapshot.ActiveWorksheetIndex = activeWorksheetIndex.Value;
                    snapshot.ActiveWorksheetName = sheetElements[activeWorksheetIndex.Value].Name?.Value;
                }

                var threadedCommentPeople = ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);

                for (int sheetIndex = 0; sheetIndex < sheetElements.Count; sheetIndex++) {
                    var sheet = sheetElements[sheetIndex];
                    if (sheet.Id == null || workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) {
                        continue;
                    }

                    var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                    var sheetName = sheet.Name?.Value ?? $"Sheet{sheetIndex + 1}";
                    IReadOnlyDictionary<string, string> resolvedFormulaTexts = this[sheetName].BuildResolvedFormulaTextMap();
                    var readerSheet = reader.GetSheet(sheetName);
                    var typedValues = BuildTypedCellMap(readerSheet);
                    string usedRangeA1 = readerSheet.GetUsedRangeA1();
                    var hyperlinkMap = ExcelWorksheetHyperlinkResolver.BuildMap(worksheetPart, usedRangeA1);
                    var commentMap = ExcelWorksheetCommentResolver.BuildLegacyCommentMap(worksheetPart);
                    var threadedCommentMap = ExcelWorksheetCommentResolver.BuildThreadedCommentMap(worksheetPart, threadedCommentPeople, sheetName);

                    var outlineProperties = worksheet.GetFirstChild<SheetProperties>()?.GetFirstChild<OutlineProperties>();
                    var sheetView = worksheet
                        .GetFirstChild<SheetViews>()?
                        .Elements<SheetView>()
                        .FirstOrDefault();
                    var worksheetSnapshot = new ExcelWorksheetSnapshot {
                        Name = sheetName,
                        Index = sheetIndex,
                        Hidden = sheet.State?.Value == SheetStateValues.Hidden || sheet.State?.Value == SheetStateValues.VeryHidden,
                        IsActive = activeWorksheetIndex == sheetIndex,
                        RightToLeft = sheetView?.RightToLeft?.Value == true,
                        ShowGridlines = sheetView?.ShowGridLines?.Value ?? true,
                        View = sheetView?.View?.InnerText,
                        ZoomScale = sheetView?.ZoomScale?.Value,
                        ZoomScaleNormal = sheetView?.ZoomScaleNormal?.Value,
                        TabColorArgb = ExcelThemeColorResolver.Resolve(worksheet.GetFirstChild<SheetProperties>()?.TabColor, workbookPart),
                        OutlineSummaryBelow = outlineProperties?.SummaryBelow?.Value,
                        OutlineSummaryRight = outlineProperties?.SummaryRight?.Value,
                        UsedRangeA1 = usedRangeA1,
                    };

                    var pane = sheetView?.GetFirstChild<Pane>();

                    if (pane != null && (pane.State?.Value == PaneStateValues.Frozen || pane.State?.Value == PaneStateValues.FrozenSplit)) {
                        worksheetSnapshot.FrozenRowCount = ConvertToInt(pane.VerticalSplit);
                        worksheetSnapshot.FrozenColumnCount = ConvertToInt(pane.HorizontalSplit);
                    }

                    var columns = worksheet.GetFirstChild<Columns>();
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
                                StyleIndex = column.Style?.Value,
                                OutlineLevel = column.OutlineLevel?.Value,
                                Collapsed = column.Collapsed?.Value == true,
                            });
                        }
                    }

                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        foreach (var row in sheetData.Elements<Row>()) {
                            var rowIndex = checked((int)(row.RowIndex?.Value ?? 0U));
                            bool customFormat = row.CustomFormat?.Value == true;
                            uint? styleIndex = row.StyleIndex?.Value;
                            byte? outlineLevel = row.OutlineLevel?.Value;
                            bool collapsed = row.Collapsed?.Value == true;
                            if (rowIndex > 0 && (row.Hidden?.Value == true || row.CustomHeight?.Value == true || row.Height != null || customFormat || styleIndex != null || outlineLevel != null || collapsed)) {
                                worksheetSnapshot.AddRow(new ExcelRowSnapshot {
                                    Index = rowIndex,
                                    Height = row.Height?.Value,
                                    Hidden = row.Hidden?.Value == true,
                                    CustomHeight = row.CustomHeight?.Value == true,
                                    CustomFormat = customFormat,
                                    StyleIndex = styleIndex,
                                    OutlineLevel = outlineLevel,
                                    Collapsed = collapsed,
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
                                    Formula = resolvedFormulaTexts.TryGetValue(cellReference, out string? formulaText)
                                        ? formulaText
                                        : cell.CellFormula?.Text,
                                    StyleIndex = cell.StyleIndex?.Value,
                                    Style = BuildCellStyleSnapshot(styleContext, cell.StyleIndex?.Value),
                                    Hyperlink = hyperlinkMap.TryGetValue(cellReference, out var hyperlink) ? hyperlink : null,
                                    Comment = commentMap.TryGetValue(cellReference, out var comment) ? comment : null,
                                    ThreadedComment = threadedCommentMap.TryGetValue(cellReference, out var threadedComments) ? threadedComments[0] : null,
                                    RichTextRuns = BuildRichTextRuns(cell, sharedStringItems),
                                });
                            }
                        }
                    }

                    AddCommentOnlyCells(worksheetSnapshot, commentMap);

                    foreach (var threadedComments in threadedCommentMap.Values) {
                        foreach (var threadedComment in threadedComments) {
                            worksheetSnapshot.AddThreadedComment(threadedComment);
                        }
                    }

                    var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
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
                        worksheet.Elements<AutoFilter>().FirstOrDefault());
                    worksheetSnapshot.Protection = BuildWorksheetProtectionSnapshot(
                        worksheet.Elements<SheetProtection>().FirstOrDefault());
                    foreach (var validation in BuildDataValidationSnapshots(worksheet.GetFirstChild<DataValidations>())) {
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

        private static IReadOnlyList<ExcelRichTextRun> BuildRichTextRuns(
            Cell cell,
            IReadOnlyList<SharedStringItem>? sharedStringItems) {
            IEnumerable<Run> runs;
            if (cell.InlineString != null) {
                runs = cell.InlineString.Elements<Run>();
            } else if (cell.DataType?.Value == CellValues.SharedString
                && int.TryParse(cell.CellValue?.InnerText, out int sharedStringIndex)
                && sharedStringIndex >= 0
                && sharedStringIndex < (sharedStringItems?.Count ?? 0)) {
                runs = sharedStringItems![sharedStringIndex].Elements<Run>();
            } else {
                return Array.Empty<ExcelRichTextRun>();
            }

            var result = new List<ExcelRichTextRun>();
            foreach (Run run in runs) {
                RunProperties? properties = run.RunProperties;
                result.Add(new ExcelRichTextRun(run.Text?.Text ?? string.Empty) {
                    Bold = properties?.GetFirstChild<Bold>() != null,
                    Italic = properties?.GetFirstChild<Italic>() != null,
                    Underline = properties?.GetFirstChild<Underline>() != null,
                    Strikethrough = properties?.GetFirstChild<Strike>() != null,
                    UnderlineStyle = ExcelRichTextRun.GetUnderlineStyle(properties),
                    FontColor = properties?.GetFirstChild<Color>()?.Rgb?.Value,
                    FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                    FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value,
                    VerticalTextAlignment = ExcelRichTextRun.GetVerticalTextAlignment(properties),
                    Outline = properties?.GetFirstChild<Outline>() != null,
                    Shadow = properties?.GetFirstChild<Shadow>() != null,
                    Condense = properties?.GetFirstChild<Condense>() != null,
                    Extend = properties?.GetFirstChild<Extend>() != null,
                    FontFamily = ExcelRichTextRun.GetFontFamily(properties),
                    FontCharacterSet = ExcelRichTextRun.GetFontCharacterSet(properties),
                });
            }

            return result;
        }

        private static int CountPackagePartsByContentType(OpenXmlPartContainer container, string marker) {
            if (string.IsNullOrWhiteSpace(marker)) return 0;

            return CountPackageParts(
                container,
                part => part.ContentType.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static int CountPackageParts(OpenXmlPartContainer container, Func<OpenXmlPart, bool> predicate) {
            if (predicate == null) throw new ArgumentNullException(nameof(predicate));

            return CountPackageParts(container, predicate, new HashSet<Uri>());
        }

        private static int CountPackageParts(
            OpenXmlPartContainer container,
            Func<OpenXmlPart, bool> predicate,
            HashSet<Uri> visitedParts) {
            int count = 0;
            foreach (var relationship in container.Parts) {
                var part = relationship.OpenXmlPart;
                if (!visitedParts.Add(part.Uri)) {
                    continue;
                }

                if (predicate(part)) {
                    count++;
                }

                count += CountPackageParts(part, predicate, visitedParts);
            }

            return count;
        }

        private static bool IsNativeSlicerPackagePart(OpenXmlPart part) {
            return string.Equals(part.ContentType, "application/vnd.ms-excel.slicer+xml", StringComparison.OrdinalIgnoreCase)
                || (string.Equals(part.ContentType, MicrosoftWorkbookSlicerCacheContentType, StringComparison.OrdinalIgnoreCase)
                    && !IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Slicer));
        }

        private static bool IsNativeTimelinePackagePart(OpenXmlPart part) {
            return string.Equals(part.ContentType, "application/vnd.ms-excel.timeline+xml", StringComparison.OrdinalIgnoreCase)
                || (string.Equals(part.ContentType, MicrosoftWorkbookTimelineCacheContentType, StringComparison.OrdinalIgnoreCase)
                    && !IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Timeline));
        }

        private static bool IsSlicerBindingMetadataPart(OpenXmlPart part) {
            return string.Equals(part.ContentType, WorkbookSlicerCacheContentType, StringComparison.OrdinalIgnoreCase)
                || IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Slicer);
        }

        private static bool IsTimelineBindingMetadataPart(OpenXmlPart part) {
            return string.Equals(part.ContentType, WorkbookTimelineCacheContentType, StringComparison.OrdinalIgnoreCase)
                || IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Timeline);
        }

        private static ExcelWorksheetProtectionSnapshot? BuildWorksheetProtectionSnapshot(SheetProtection? protection) {
            if (protection == null) {
                return null;
            }

            return new ExcelWorksheetProtectionSnapshot {
                AllowSelectLockedCells = IsProtectionActionAllowed(protection.SelectLockedCells, lockedWhenOmitted: false),
                AllowSelectUnlockedCells = IsProtectionActionAllowed(protection.SelectUnlockedCells, lockedWhenOmitted: false),
                AllowFormatCells = IsProtectionActionAllowed(protection.FormatCells, lockedWhenOmitted: true),
                AllowFormatColumns = IsProtectionActionAllowed(protection.FormatColumns, lockedWhenOmitted: true),
                AllowFormatRows = IsProtectionActionAllowed(protection.FormatRows, lockedWhenOmitted: true),
                AllowInsertColumns = IsProtectionActionAllowed(protection.InsertColumns, lockedWhenOmitted: true),
                AllowInsertRows = IsProtectionActionAllowed(protection.InsertRows, lockedWhenOmitted: true),
                AllowInsertHyperlinks = IsProtectionActionAllowed(protection.InsertHyperlinks, lockedWhenOmitted: true),
                AllowDeleteColumns = IsProtectionActionAllowed(protection.DeleteColumns, lockedWhenOmitted: true),
                AllowDeleteRows = IsProtectionActionAllowed(protection.DeleteRows, lockedWhenOmitted: true),
                AllowSort = IsProtectionActionAllowed(protection.Sort, lockedWhenOmitted: true),
                AllowAutoFilter = IsProtectionActionAllowed(protection.AutoFilter, lockedWhenOmitted: true),
                AllowPivotTables = IsProtectionActionAllowed(protection.PivotTables, lockedWhenOmitted: true),
            };
        }

        private static bool IsProtectionActionAllowed(BooleanValue? lockFlag, bool lockedWhenOmitted) {
            return !(lockFlag?.Value ?? lockedWhenOmitted);
        }

        private static void AddCommentOnlyCells(ExcelWorksheetSnapshot worksheetSnapshot, IReadOnlyDictionary<string, ExcelCommentSnapshot> commentMap) {
            if (commentMap.Count == 0) {
                return;
            }

            var existingCells = new HashSet<string>(
                worksheetSnapshot.Cells.Select(cell => A1.CellReference(cell.Row, cell.Column)),
                StringComparer.OrdinalIgnoreCase);

            foreach (var pair in commentMap) {
                if (existingCells.Contains(pair.Key)) {
                    continue;
                }

                var (row, column) = A1.ParseCellRef(pair.Key);
                if (row <= 0 || column <= 0) {
                    continue;
                }

                worksheetSnapshot.AddCell(new ExcelCellSnapshot {
                    Row = row,
                    Column = column,
                    Comment = pair.Value,
                });
            }
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
            bool hasSimpleGradient = context.TryGetSimpleGradientFill(fill, out ExcelGradientFillInfo gradient);

            return new ExcelCellStyleSnapshot {
                StyleIndex = styleIndex.Value,
                NumberFormatId = numberFormatId,
                NumberFormatCode = context.GetNumberFormatCode(numberFormatId),
                IsDateLike = context.IsDateLike(numberFormatId),
                Bold = font?.Bold != null,
                Italic = font?.Italic != null,
                Underline = font?.Underline != null,
                Strikethrough = font?.Strike != null,
                FontName = font?.FontName?.Val?.Value,
                FontSize = font?.FontSize?.Val?.Value,
                FontColorArgb = context.GetColorArgb(font?.Color),
                FillColorArgb = context.GetFillColorArgb(fill),
                FillPatternType = context.GetFillPatternType(fill),
                FillPatternForegroundColorArgb = context.GetFillPatternForegroundColorArgb(fill),
                FillPatternBackgroundColorArgb = context.GetFillPatternBackgroundColorArgb(fill),
                FillGradientUnsupported = fill?.GradientFill != null && !hasSimpleGradient,
                FillGradientStartColorArgb = hasSimpleGradient ? gradient.StartColorArgb : null,
                FillGradientEndColorArgb = hasSimpleGradient ? gradient.EndColorArgb : null,
                FillGradientStops = hasSimpleGradient ? CreateGradientStopSnapshots(gradient) : Array.Empty<ExcelGradientFillStopSnapshot>(),
                FillGradientDegree = hasSimpleGradient ? gradient.Degree : null,
                Border = BuildBorderSnapshot(context, border),
                HorizontalAlignment = cellFormat.Alignment?.Horizontal?.InnerText,
                VerticalAlignment = cellFormat.Alignment?.Vertical?.InnerText,
                TextRotation = ToTextRotation(cellFormat.Alignment?.TextRotation?.Value),
                TextIndent = cellFormat.Alignment?.Indent?.Value,
                WrapText = cellFormat.Alignment?.WrapText?.Value == true,
                ShrinkToFit = cellFormat.Alignment?.ShrinkToFit?.Value == true,
            };
        }

        private static int? ToTextRotation(uint? value) {
            if (!value.HasValue) {
                return null;
            }

            return value.Value <= int.MaxValue ? (int)value.Value : (int?)null;
        }

        private static IReadOnlyList<ExcelGradientFillStopSnapshot> CreateGradientStopSnapshots(ExcelGradientFillInfo gradient) =>
            gradient.Stops.Select(stop => new ExcelGradientFillStopSnapshot(stop.Offset, stop.ColorArgb)).ToArray();

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

        private static ExcelCellBorderSnapshot? BuildBorderSnapshot(StyleInspectionContext context, Border? border) {
            if (border == null) {
                return null;
            }

            var left = BuildBorderSideSnapshot(context, border.LeftBorder);
            var right = BuildBorderSideSnapshot(context, border.RightBorder);
            var top = BuildBorderSideSnapshot(context, border.TopBorder);
            var bottom = BuildBorderSideSnapshot(context, border.BottomBorder);
            var diagonal = BuildBorderSideSnapshot(context, border.DiagonalBorder);
            bool diagonalUp = border.DiagonalUp?.Value == true;
            bool diagonalDown = border.DiagonalDown?.Value == true;

            if (left == null && right == null && top == null && bottom == null && (!diagonalUp && !diagonalDown || diagonal == null)) {
                return null;
            }

            return new ExcelCellBorderSnapshot {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Diagonal = diagonal,
                DiagonalUp = diagonalUp,
                DiagonalDown = diagonalDown,
            };
        }

        private static ExcelBorderSideSnapshot? BuildBorderSideSnapshot(StyleInspectionContext context, BorderPropertiesType? borderSide) {
            if (borderSide == null) {
                return null;
            }

            var style = ExtractBorderStyle(borderSide);

            var colorArgb = context.GetColorArgb(borderSide.GetFirstChild<Color>());

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

        private static int? GetActiveWorksheetIndex(Workbook workbook, int sheetCount) {
            if (sheetCount <= 0) {
                return null;
            }

            uint activeTab = workbook.GetFirstChild<BookViews>()?
                .Elements<WorkbookView>()
                .FirstOrDefault()?
                .ActiveTab?.Value ?? 0U;

            if (activeTab >= sheetCount) {
                return sheetCount - 1;
            }

            return checked((int)activeTab);
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
            private readonly WorkbookPart? _workbookPart;

            private StyleInspectionContext(
                WorkbookPart? workbookPart,
                IReadOnlyList<CellFormat> cellFormats,
                IReadOnlyList<Font> fonts,
                IReadOnlyList<Fill> fills,
                IReadOnlyList<Border> borders,
                Dictionary<uint, string> numberFormats) {
                _workbookPart = workbookPart;
                _cellFormats = cellFormats;
                _fonts = fonts;
                _fills = fills;
                _borders = borders;
                _numberFormats = numberFormats;
            }

            internal static StyleInspectionContext Create(WorkbookPart? workbookPart, Stylesheet? stylesheet) {
                var numberFormats = new Dictionary<uint, string>();
                if (stylesheet?.NumberingFormats != null) {
                    foreach (var numberingFormat in stylesheet.NumberingFormats.Elements<NumberingFormat>()) {
                        if (numberingFormat.NumberFormatId?.Value is uint id && numberingFormat.FormatCode?.Value is string code) {
                            numberFormats[id] = code;
                        }
                    }
                }

                return new StyleInspectionContext(
                    workbookPart,
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
                    || (_numberFormats.TryGetValue(numberFormatId, out var code) && ExcelNumberFormatClassifier.LooksLikeDateFormat(code));
            }

            internal string? GetColorArgb(OpenXmlElement? colorElement) {
                return ExcelThemeColorResolver.Resolve(colorElement, _workbookPart);
            }

            internal string? GetFillColorArgb(Fill? fill) {
                var patternFill = fill?.PatternFill;
                if (patternFill == null) {
                    return null;
                }

                if (patternFill.PatternType?.Value == PatternValues.Solid) {
                    return GetColorArgb(patternFill.ForegroundColor) ?? GetColorArgb(patternFill.BackgroundColor);
                }

                return GetColorArgb(patternFill.BackgroundColor);
            }

            internal string? GetFillPatternType(Fill? fill) {
                var patternFill = fill?.PatternFill;
                if (patternFill?.PatternType?.Value == null) {
                    return fill?.GradientFill != null ? "gradient" : null;
                }

                if (patternFill.PatternType.Value == PatternValues.None) {
                    return null;
                }

                string text = patternFill.PatternType.InnerText ?? string.Empty;
                return string.IsNullOrEmpty(text) ? string.Empty : char.ToLowerInvariant(text[0]) + text.Substring(1);
            }

            internal string? GetFillPatternForegroundColorArgb(Fill? fill) {
                var patternFill = fill?.PatternFill;
                return patternFill == null ? null : GetColorArgb(patternFill.ForegroundColor);
            }

            internal string? GetFillPatternBackgroundColorArgb(Fill? fill) {
                var patternFill = fill?.PatternFill;
                return patternFill == null ? null : GetColorArgb(patternFill.BackgroundColor);
            }

            internal bool TryGetSimpleGradientFill(Fill? fill, out ExcelGradientFillInfo gradient) =>
                ExcelGradientFillResolver.TryResolveSimpleLinearGradient(fill, _workbookPart, out gradient);

            private static bool IsBuiltInDate(uint id) {
                return id is 14 or 15 or 16 or 17 or 18 or 19 or 20 or 21 or 22
                    or 27 or 30 or 36 or 45 or 46 or 47;
            }

        }
    }
}
