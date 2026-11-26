using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Whether the workbook has a worksheet. Unlike <see cref="Sheets"/>, it leaves rows waiting for the direct writer
        /// out of the worksheets, so a save that follows keeps them on that writer.
        /// </summary>
        internal bool HasWorksheets => ExecuteReadWithoutMaterializing(() =>
            _deferredWorkbook is { } deferred ? deferred.Sheets.Any(sheet => sheet.IsWorksheet)
            : _directDataSetSaveCandidate?.Model.Sheets.Count > 0
            || _materializedDirectDataSetFastSaveModel?.Sheets.Count > 0
            || ReadSheetElements().Any(sheet => sheet.Id?.Value is string id && WorkbookPartRoot.GetPartById(id) is WorksheetPart));

        /// <summary>Counts all workbook sheet entries, including chart sheets, for source-level inspection.</summary>
        internal int GetWorkbookSheetCount(CancellationToken cancellationToken) {
            MaterializeDeferredDataSetImport(cancellationToken);
            return Locking.ExecuteRead(EnsureLock(), () => WorkbookRoot.Sheets?.Elements<Sheet>().Count() ?? 0);
        }

        /// <summary>
        /// Creates an immutable workbook snapshot that exposes OfficeIMO's current workbook model for downstream integrations.
        /// </summary>
        public ExcelWorkbookInfo Inspect(ExcelReaderOptions? options = null) => Inspect(options, CancellationToken.None);

        /// <summary>Creates the native immutable workbook snapshot with cancellation between source sheets, rows and cells.</summary>
        public ExcelWorkbookInfo Inspect(ExcelReaderOptions? options, CancellationToken cancellationToken) =>
            Inspect(options, cancellationToken, includeCellPresentation: true);

        /// <summary>
        /// Builds the workbook snapshot. Without cell presentation, cells omit their style and rich-text run snapshots,
        /// which text extraction does not use. With <paramref name="selectedSheet"/>, only worksheets of that name
        /// (ignoring case) are inspected; workbook-level facts still cover the whole workbook.
        /// </summary>
        internal ExcelWorkbookInfo Inspect(ExcelReaderOptions? options, CancellationToken cancellationToken, bool includeCellPresentation, string? selectedSheet = null) {
            string? selectedSheetName = string.IsNullOrWhiteSpace(selectedSheet) ? null : selectedSheet!.Trim();
            cancellationToken.ThrowIfCancellationRequested();
            var effectiveOptions = options ?? new ExcelReaderOptions {
                UseCachedFormulaResult = true,
                TreatDatesUsingNumberFormat = true,
            };

            var snapshot = new ExcelWorkbookInfo {
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
                // Style snapshots are part of cell presentation; without it the stylesheet DOM is not needed.
                var styleContext = includeCellPresentation
                    ? StyleInspectionContext.Create(workbookPart, workbookPart.WorkbookStylesPart?.Stylesheet)
                    : null;
                // Shared-string items are only needed for rich-text runs; loading them parses the whole table.
                var sharedStringTablePart = includeCellPresentation ? workbookPart.SharedStringTablePart : null;
                var sharedStringItems = sharedStringTablePart != null && SharedStringTableMayHaveRuns(sharedStringTablePart)
                    ? sharedStringTablePart.SharedStringTable?.Elements<SharedStringItem>().ToList()
                    : null;
                CountInspectedPackageParts(workbookPart, snapshot);
                var sheetElements = workbook.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                int? activeWorksheetIndex = GetActiveWorksheetIndex(workbook, sheetElements.Count);
                if (activeWorksheetIndex.HasValue) {
                    snapshot.ActiveWorksheetIndex = activeWorksheetIndex.Value;
                    snapshot.ActiveWorksheetName = sheetElements[activeWorksheetIndex.Value].Name?.Value;
                }

                var threadedCommentPeople = ExcelWorksheetCommentResolver.BuildThreadedCommentPersonMap(workbookPart);

                for (int sheetIndex = 0; sheetIndex < sheetElements.Count; sheetIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var sheet = sheetElements[sheetIndex];
                    if (sheet.Id == null || workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) {
                        continue;
                    }

                    var sheetName = sheet.Name?.Value ?? $"Sheet{sheetIndex + 1}";
                    if (selectedSheetName != null && !string.Equals(sheetName, selectedSheetName, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    var worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                    ExcelSheet sourceSheet = this[sheetName];
                    IReadOnlyDictionary<string, string> resolvedFormulaTexts = sourceSheet.BuildResolvedFormulaTextMap();
                    var readerSheet = reader.GetSheet(sheetName, worksheetPart);
                    string usedRangeA1 = readerSheet.GetUsedRangeA1();
                    var hyperlinkMap = ExcelWorksheetHyperlinkResolver.BuildMap(worksheetPart, usedRangeA1);
                    var commentMap = ExcelWorksheetCommentResolver.BuildLegacyCommentMap(worksheetPart);
                    var threadedCommentMap = ExcelWorksheetCommentResolver.BuildThreadedCommentMap(worksheetPart, threadedCommentPeople, sheetName);

                    var outlineProperties = worksheet.GetFirstChild<SheetProperties>()?.GetFirstChild<OutlineProperties>();
                    var sheetView = worksheet
                        .GetFirstChild<SheetViews>()?
                        .Elements<SheetView>()
                        .FirstOrDefault();
                    var worksheetSnapshot = new ExcelWorksheetInfo {
                        Name = sheetName,
                        ConditionalFormattingRuleCount = sourceSheet.CountConditionalFormattingRules(),
                        Index = sheetIndex,
                        Hidden = sheet.State?.Value == SheetStateValues.Hidden || sheet.State?.Value == SheetStateValues.VeryHidden,
                        VisibilityState = sheet.State?.Value == SheetStateValues.VeryHidden ? "veryHidden" :
                            sheet.State?.Value == SheetStateValues.Hidden ? "hidden" : "visible",
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

                    foreach (ExcelColumnInfo column in sourceSheet.GetColumnDefinitions()) {
                        worksheetSnapshot.AddColumn(column);
                    }

                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        foreach (var row in sheetData.Elements<Row>()) {
                            cancellationToken.ThrowIfCancellationRequested();
                            var rowIndex = checked((int)(row.RowIndex?.Value ?? 0U));
                            bool customFormat = row.CustomFormat?.Value == true;
                            uint? styleIndex = row.StyleIndex?.Value;
                            byte? outlineLevel = row.OutlineLevel?.Value;
                            bool collapsed = row.Collapsed?.Value == true;
                            if (rowIndex > 0 && (row.Hidden?.Value == true || row.CustomHeight?.Value == true || row.Height != null || customFormat || styleIndex != null || outlineLevel != null || collapsed)) {
                                worksheetSnapshot.AddRow(new ExcelRowInfo {
                                    Index = rowIndex,
                                    StoredHeight = row.Height?.Value,
                                    Hidden = row.Hidden?.Value == true,
                                    CustomHeight = row.CustomHeight?.Value == true,
                                    CustomFormat = customFormat,
                                    StyleIndex = styleIndex,
                                    OutlineLevel = outlineLevel,
                                    Collapsed = collapsed,
                                });
                            }

                            foreach (var cell in row.Elements<Cell>()) {
                                cancellationToken.ThrowIfCancellationRequested();
                                if (string.IsNullOrWhiteSpace(cell.CellReference?.Value)) {
                                    continue;
                                }

                                var cellReference = cell.CellReference!.Value!;
                                var (_, columnIndex) = ExcelA1.ParseCellRef(cellReference);
                                if (rowIndex <= 0 || columnIndex <= 0) {
                                    continue;
                                }

                                // One walk over the cell's children; each SDK child property searches them again.
                                ExcelSheetReader.GetCellChildren(cell, out CellFormula? cellFormula, out CellValue? cellValue, out InlineString? inlineString);
                                string? storedFormula = cellFormula?.Text;
                                uint? cellStyleIndex = cell.StyleIndex?.Value;
                                worksheetSnapshot.AddCell(new ExcelCellInfo {
                                    Row = rowIndex,
                                    Column = columnIndex,
                                    // The worksheet DOM is already loaded, so values come from its cells instead of another read of the part.
                                    Value = readerSheet.ConvertLoadedCell(cell),
                                    Formula = resolvedFormulaTexts.TryGetValue(cellReference, out string? formulaText)
                                        ? formulaText
                                        : storedFormula,
                                    StoredFormula = storedFormula,
                                    FormulaKind = cellFormula == null ? null : cellFormula.FormulaType?.InnerText ?? "normal",
                                    HasStoredValue = cellValue != null,
                                    StoredValue = cellValue?.Text,
                                    StoredValueType = cell.DataType?.InnerText ?? "n",
                                    StyleIndex = cellStyleIndex,
                                    Style = styleContext != null ? BuildCellStyleSnapshot(styleContext, cellStyleIndex) : null,
                                    Hyperlink = hyperlinkMap.TryGetValue(cellReference, out var hyperlink) ? hyperlink : null,
                                    Comment = commentMap.TryGetValue(cellReference, out var comment) ? comment : null,
                                    ThreadedComment = threadedCommentMap.TryGetValue(cellReference, out var threadedComments) ? threadedComments[0] : null,
                                    RichTextRuns = includeCellPresentation ? BuildRichTextRuns(cell, cellValue, inlineString, sharedStringItems) : Array.Empty<ExcelRichTextRun>(),
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
                            if (!ExcelA1.TryParseRange(reference!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                                continue;
                            }

                            worksheetSnapshot.AddMergedRange(new ExcelMergedRangeInfo {
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

                    snapshot.AddSheet(worksheetSnapshot);
                }

                var definedNames = workbook.DefinedNames;
                if (definedNames != null) {
                    foreach (var definedName in definedNames.Elements<DefinedName>()) {
                        var name = definedName.Name?.Value;
                        if (string.IsNullOrWhiteSpace(name)) {
                            continue;
                        }

                        snapshot.AddNamedRange(new ExcelNamedRangeInspectionInfo {
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
            CellValue? cellValue,
            InlineString? inlineString,
            IReadOnlyList<SharedStringItem>? sharedStringItems) {
            IEnumerable<Run> runs;
            if (inlineString != null) {
                runs = inlineString.Elements<Run>();
            } else if (cell.DataType?.Value == CellValues.SharedString
                && int.TryParse(cellValue?.InnerText, out int sharedStringIndex)
                && sharedStringIndex >= 0
                && sharedStringIndex < (sharedStringItems?.Count ?? 0)) {
                runs = sharedStringItems![sharedStringIndex].Elements<Run>();
            } else {
                return Array.Empty<ExcelRichTextRun>();
            }

            var result = new List<ExcelRichTextRun>();
            foreach (Run run in runs) {
                RunProperties? properties = run.RunProperties;
                result.Add(ExcelRichTextRun.FromOpenXml(run.Text?.Text ?? string.Empty, properties));
            }

            return result;
        }

        /// <summary>
        /// Without runs in the table, no cell has shared-string rich text, so its DOM is not needed. A streamed scan
        /// answers that for an unloaded part read as the package would read it; a loaded part, markup compatibility
        /// processing, or XML the scan cannot read keeps the DOM, which also reports a part over the size limit.
        /// </summary>
        private static bool SharedStringTableMayHaveRuns(SharedStringTablePart part) {
            if (part.IsRootElementLoaded
                || part.OpenXmlPackage.MarkupCompatibilityProcessSettings.ProcessMode != MarkupCompatibilityProcessMode.NoProcess) {
                return true;
            }

            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                CloseInput = false
            };
            long maxCharacters = part.OpenXmlPackage.MaxCharactersInPart;
            if (maxCharacters > 0) {
                settings.MaxCharactersInDocument = maxCharacters;
            }

            try {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                using XmlReader reader = XmlReader.Create(stream, settings);
                while (reader.Read()) {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "r") {
                        return true;
                    }
                }

                return false;
            } catch (XmlException) {
                return true;
            }
        }

        /// <summary>Counts the package parts the snapshot reports in one walk of the part graph.</summary>
        private static void CountInspectedPackageParts(WorkbookPart workbookPart, ExcelWorkbookInfo snapshot) {
            var counts = new InspectedPartCounts();
            CountInspectedPackageParts(workbookPart, ref counts, new HashSet<Uri>());
            snapshot.SlicerPartCount = counts.Slicer;
            snapshot.TimelinePartCount = counts.Timeline;
            snapshot.SlicerBindingMetadataPartCount = counts.SlicerBindingMetadata;
            snapshot.TimelineBindingMetadataPartCount = counts.TimelineBindingMetadata;
            snapshot.ConnectionPartCount = counts.Connection;
            snapshot.QueryTablePartCount = counts.QueryTable;
            snapshot.ChartPartCount = counts.Chart;
            snapshot.PivotTablePartCount = counts.PivotTable;
        }

        private static void CountInspectedPackageParts(
            OpenXmlPartContainer container,
            ref InspectedPartCounts counts,
            HashSet<Uri> visitedParts) {
            foreach (var relationship in container.Parts) {
                var part = relationship.OpenXmlPart;
                if (!visitedParts.Add(part.Uri)) {
                    continue;
                }

                if (IsNativeSlicerPackagePart(part)) counts.Slicer++;
                if (IsNativeTimelinePackagePart(part)) counts.Timeline++;
                if (IsSlicerBindingMetadataPart(part)) counts.SlicerBindingMetadata++;
                if (IsTimelineBindingMetadataPart(part)) counts.TimelineBindingMetadata++;
                if (part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0) counts.Connection++;
                if (part.ContentType.IndexOf("queryTable", StringComparison.OrdinalIgnoreCase) >= 0) counts.QueryTable++;
                if (part is ChartPart || part is ExtendedChartPart) counts.Chart++;
                if (part is PivotTablePart) counts.PivotTable++;

                CountInspectedPackageParts(part, ref counts, visitedParts);
            }
        }

        private struct InspectedPartCounts {
            internal int Slicer;
            internal int Timeline;
            internal int SlicerBindingMetadata;
            internal int TimelineBindingMetadata;
            internal int Connection;
            internal int QueryTable;
            internal int Chart;
            internal int PivotTable;
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
            return CombinedPivotInteractionMetadataContains(part, ExcelPivotInteractionCacheKind.Slicer)
                || string.Equals(part.ContentType, WorkbookSlicerCacheContentType, StringComparison.OrdinalIgnoreCase)
                || IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Slicer);
        }

        private static bool IsTimelineBindingMetadataPart(OpenXmlPart part) {
            return CombinedPivotInteractionMetadataContains(part, ExcelPivotInteractionCacheKind.Timeline)
                || string.Equals(part.ContentType, WorkbookTimelineCacheContentType, StringComparison.OrdinalIgnoreCase)
                || IsLegacyOfficeImoPivotInteractionMetadataPart(part, ExcelPivotInteractionCacheKind.Timeline);
        }

        private static ExcelWorksheetProtectionInfo? BuildWorksheetProtectionSnapshot(SheetProtection? protection) {
            if (protection == null) {
                return null;
            }

            return new ExcelWorksheetProtectionInfo {
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

        private static void AddCommentOnlyCells(ExcelWorksheetInfo worksheetSnapshot, IReadOnlyDictionary<string, ExcelCommentInspectionInfo> commentMap) {
            if (commentMap.Count == 0) {
                return;
            }

            var existingCells = new HashSet<string>(
                worksheetSnapshot.Cells.Select(cell => ExcelA1.CellReference(cell.Row, cell.Column)),
                StringComparer.OrdinalIgnoreCase);

            foreach (var pair in commentMap) {
                if (existingCells.Contains(pair.Key)) {
                    continue;
                }

                var (row, column) = ExcelA1.ParseCellRef(pair.Key);
                if (row <= 0 || column <= 0) {
                    continue;
                }

                worksheetSnapshot.AddCell(new ExcelCellInfo {
                    Row = row,
                    Column = column,
                    Comment = pair.Value,
                });
            }
        }
        /// <summary>Returns the style snapshot for a cell format, built once per style index and shared by its cells.</summary>
        private static ExcelCellStyleInfo? BuildCellStyleSnapshot(StyleInspectionContext context, uint? styleIndex) {
            if (!styleIndex.HasValue) {
                return null;
            }

            return context.GetOrAddStyleSnapshot(styleIndex.Value, static (ctx, index) => BuildCellStyleSnapshotCore(ctx, index));
        }

        private static ExcelCellStyleInfo? BuildCellStyleSnapshotCore(StyleInspectionContext context, uint styleIndex) {
            var cellFormat = context.GetCellFormat(styleIndex);
            if (cellFormat == null) {
                return null;
            }

            var numberFormatId = cellFormat.NumberFormatId?.Value ?? 0U;
            var font = context.GetFont(cellFormat.FontId?.Value ?? 0U);
            var fill = context.GetFill(cellFormat.FillId?.Value ?? 0U);
            var border = context.GetBorder(cellFormat.BorderId?.Value ?? 0U);
            bool hasSimpleGradient = context.TryGetSimpleGradientFill(fill, out ExcelGradientFillInfo gradient);

            return new ExcelCellStyleInfo {
                StyleIndex = styleIndex,
                NumberFormatId = numberFormatId,
                NumberFormatCode = context.GetNumberFormatCode(numberFormatId),
                IsDateLike = context.IsDateLike(numberFormatId),
                Bold = ExcelOpenXmlFontProperty.IsEnabled(font?.Bold),
                Italic = ExcelOpenXmlFontProperty.IsEnabled(font?.Italic),
                Underline = ExcelOpenXmlFontProperty.IsUnderlineEnabled(font?.Underline),
                UnderlineStyle = font?.Underline == null
                    ? null
                    : (font.Underline.Val?.Value ?? UnderlineValues.Single).ToOfficeEnum(),
                Strikethrough = ExcelOpenXmlFontProperty.IsEnabled(font?.Strike),
                VerticalTextAlignment = font?.VerticalTextAlignment?.Val?.Value.ToOfficeEnum(),
                FontName = font?.FontName?.Val?.Value,
                IsFontFamilyExplicit = (cellFormat.FontId?.Value ?? 0U) != 0U,
                FontSize = font?.FontSize?.Val?.Value is double fontSize ? OfficeLength.Points(fontSize) : null,
                FontColorArgb = context.GetColorArgb(font?.Color),
                FillColorArgb = context.GetFillColorArgb(fill),
                FillPatternType = context.GetFillPatternType(fill),
                FillPatternForegroundColorArgb = context.GetFillPatternForegroundColorArgb(fill),
                FillPatternBackgroundColorArgb = context.GetFillPatternBackgroundColorArgb(fill),
                FillGradientUnsupported = fill?.GradientFill != null && !hasSimpleGradient,
                FillGradientStartColorArgb = hasSimpleGradient ? gradient.StartColorArgb : null,
                FillGradientEndColorArgb = hasSimpleGradient ? gradient.EndColorArgb : null,
                FillGradientStops = hasSimpleGradient ? CreateGradientStopSnapshots(gradient) : Array.Empty<ExcelGradientFillStopInspectionInfo>(),
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

        private static IReadOnlyList<ExcelGradientFillStopInspectionInfo> CreateGradientStopSnapshots(ExcelGradientFillInfo gradient) =>
            gradient.Stops.Select(stop => new ExcelGradientFillStopInspectionInfo(stop.Offset, stop.ColorArgb)).ToArray();

        private static ExcelAutoFilterInspectionInfo? BuildAutoFilterSnapshot(AutoFilter? autoFilter) {
            var reference = autoFilter?.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                return null;
            }

            int rowStart;
            int columnStart;
            int rowEnd;
            int columnEnd;
            if (!ExcelA1.TryParseRange(reference!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                return null;
            }

            var snapshot = new ExcelAutoFilterInspectionInfo {
                A1Range = reference!,
                StartRow = rowStart,
                EndRow = rowEnd,
                StartColumn = columnStart,
                EndColumn = columnEnd,
            };

            foreach (var filterColumn in autoFilter!.Elements<FilterColumn>()) {
                var columnId = checked((int)(filterColumn.ColumnId?.Value ?? 0U));
                var columnSnapshot = new ExcelFilterColumnInfo {
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
                    var customFiltersSnapshot = new ExcelCustomFiltersInfo {
                        MatchAll = customFilters.And?.Value == true,
                    };

                    foreach (var customFilter in customFilters.Elements<CustomFilter>()) {
                        var value = customFilter.Val?.Value;
                        if (string.IsNullOrWhiteSpace(value)) {
                            continue;
                        }

                        customFiltersSnapshot.AddCondition(new ExcelCustomFilterConditionInfo {
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

        private static IReadOnlyList<ExcelDataValidationInspectionInfo> BuildDataValidationSnapshots(DataValidations? dataValidations) {
            var snapshots = new List<ExcelDataValidationInspectionInfo>();
            if (dataValidations == null) {
                return snapshots;
            }

            foreach (var dataValidation in dataValidations.Elements<DataValidation>()) {
                var snapshot = new ExcelDataValidationInspectionInfo {
                    Type = GetOpenXmlAttributeValue(dataValidation, "type")?.ToLowerInvariant(),
                    Operator = GetOpenXmlAttributeValue(dataValidation, "operator"),
                    AllowBlank = dataValidation.AllowBlank?.Value == true,
                    SuppressDropDown = dataValidation.ShowDropDown?.Value == true,
                    Formula1 = dataValidation.GetFirstChild<Formula1>()?.Text,
                    Formula2 = dataValidation.GetFirstChild<Formula2>()?.Text,
                    ShowInputMessage = dataValidation.ShowInputMessage?.Value == true,
                    PromptTitle = dataValidation.PromptTitle?.Value,
                    Prompt = dataValidation.Prompt?.Value,
                    ShowErrorMessage = dataValidation.ShowErrorMessage?.Value == true,
                    ErrorTitle = dataValidation.ErrorTitle?.Value,
                    Error = dataValidation.Error?.Value,
                    ErrorStyle = GetOpenXmlAttributeValue(dataValidation, "errorStyle"),
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

        private static ExcelTableInspectionInfo? BuildTableSnapshot(Table? table) {
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
            if (!ExcelA1.TryParseRange(range!, out rowStart, out columnStart, out rowEnd, out columnEnd)) {
                return null;
            }

            var snapshot = new ExcelTableInspectionInfo {
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
                snapshot.AddColumn(new ExcelTableColumnInspectionInfo {
                    Index = checked((int)(tableColumn.Id?.Value ?? 0U)),
                    Name = tableColumn.Name?.Value ?? string.Empty,
                    TotalsRowFunction = GetOpenXmlAttributeValue(tableColumn, "totalsRowFunction"),
                });
            }

            return snapshot;
        }

        private static string? GetOpenXmlAttributeValue(OpenXmlElement element, string localName) {
            ArgumentNullException.ThrowIfNull(element);
            ArgumentException.ThrowIfNullOrWhiteSpace(localName);

            var attribute = element.GetAttributes().FirstOrDefault(a => string.Equals(a.LocalName, localName, StringComparison.OrdinalIgnoreCase));
            return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
        }

        private static ExcelCellBorderInfo? BuildBorderSnapshot(StyleInspectionContext context, Border? border) {
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

            return new ExcelCellBorderInfo {
                Left = left,
                Right = right,
                Top = top,
                Bottom = bottom,
                Diagonal = diagonal,
                DiagonalUp = diagonalUp,
                DiagonalDown = diagonalDown,
            };
        }

        private static ExcelBorderSideInfo? BuildBorderSideSnapshot(StyleInspectionContext context, BorderPropertiesType? borderSide) {
            if (borderSide == null) {
                return null;
            }

            var style = ExtractBorderStyle(borderSide);

            var colorArgb = context.GetColorArgb(borderSide.GetFirstChild<Color>());

            if (string.IsNullOrWhiteSpace(style) && string.IsNullOrWhiteSpace(colorArgb)) {
                return null;
            }

            return new ExcelBorderSideInfo {
                Style = style,
                ColorArgb = colorArgb,
            };
        }

        private static string? ExtractBorderStyle(BorderPropertiesType borderSide) {
            // The attribute text is kept as written, including values outside the schema enumeration.
            string? value = borderSide.Style?.InnerText;
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim().ToLowerInvariant();
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

            private readonly Dictionary<uint, ExcelCellStyleInfo?> _styleSnapshots = new();

            internal ExcelCellStyleInfo? GetOrAddStyleSnapshot(uint styleIndex, Func<StyleInspectionContext, uint, ExcelCellStyleInfo?> build) {
                if (!_styleSnapshots.TryGetValue(styleIndex, out ExcelCellStyleInfo? snapshot)) {
                    snapshot = build(this, styleIndex);
                    _styleSnapshots[styleIndex] = snapshot;
                }
                return snapshot;
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
