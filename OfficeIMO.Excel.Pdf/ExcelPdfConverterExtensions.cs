using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    /// <summary>
    /// First-party Excel workbook to PDF conversion helpers.
    /// </summary>
    public static class ExcelPdfConverterExtensions {
        /// <summary>
        /// Converts an Excel workbook to a first-party OfficeIMO PDF document model.
        /// </summary>
        public static PdfCore.PdfDoc ToPdfDocument(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            options ??= new ExcelPdfSaveOptions();
            options.Warnings.Clear();
            var pdf = PdfCore.PdfDoc.Create(CreatePdfOptions(options));
            using ExcelDocumentReader reader = document.CreateReader();
            IReadOnlyList<string> sheetNames = GetSheetNames(reader, options);
            bool hasExplicitSheetSelection = HasExplicitSheetSelection(options);
            IReadOnlyList<WorksheetPdfExportPlan> exportPlans = BuildWorksheetExportPlans(document, reader, sheetNames, options, hasExplicitSheetSelection);
            IReadOnlyDictionary<string, string> sheetDestinations = BuildSheetDestinationMap(exportPlans);
            IReadOnlyDictionary<string, string> cellDestinations = BuildCellDestinationMap(exportPlans);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                object?[,] values = plan.ExportData.Values;
                int columns = values.GetLength(1);

                pdf.Section(page => {
                    ApplyWorksheetPageSetup(page, plan.PageSetup, options);
                    ApplyWorksheetHeaderFooter(page, plan.HeaderFooter, plan.SheetName, document.FilePath, options);
                    page.Content(content => content.Item(item => {
                        item.Bookmark(plan.BookmarkName);
                        if (options.IncludeSheetHeadings) {
                            item.H1(plan.SheetName);
                        }

                        IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>> imagesByCellReference = CreateWorksheetImageMap(plan);
                        foreach (WorksheetImageExportData image in plan.Images) {
                            if (!imagesByCellReference.ContainsKey(NormalizeCellReference(image.CellReference))) {
                                item.Image(image.Bytes, image.WidthPoints, image.HeightPoints, PdfCore.PdfAlign.Left, spacingBefore: 4, spacingAfter: 6);
                            }
                        }

                        foreach (WorksheetChartExportData chart in plan.Charts) {
                            AddWorksheetChart(item, chart);
                        }

                        if (plan.HasTable) {
                            IReadOnlyList<TableChunk> chunks = CreateTableChunks(plan, options, columns);
                            for (int chunkIndex = 0; chunkIndex < chunks.Count; chunkIndex++) {
                                TableChunk chunk = chunks[chunkIndex];
                                if (chunkIndex > 0) {
                                    item.PageBreak();
                                }

                                item.Table(
                                    CreatePdfRows(values, plan.ExportData.Styles, plan.ExportData.Hyperlinks, plan.ExportData.CellReferences, plan.ExportData.MergedCells, imagesByCellReference, chunk.RowIndexes, chunk.StartColumn, chunk.ColumnCount, options.EmptyCellText, sheetDestinations, cellDestinations, plan.SheetName),
                                    style: CreateTableStyle(options, plan.PageSetup, chunk.RowIndexes, chunk.HeaderRowCount, plan.ExportData.Styles, plan.ExportData.ConditionalFills, plan.ExportData.ColumnWidths, plan.ExportData.RowHeights, chunk.StartColumn, chunk.ColumnCount));
                            }
                        }
                    }));
                });
            }

            if (exportPlans.Count == 0) {
                pdf.H1("Workbook");
                pdf.Table(new[] { new[] { "No worksheet data found." } }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 0 });
            }

            return pdf;
        }

        /// <summary>
        /// Converts an Excel workbook to PDF bytes.
        /// </summary>
        public static byte[] SaveAsPdf(this ExcelDocument document, ExcelPdfSaveOptions? options = null) {
            return document.ToPdfDocument(options).ToBytes();
        }

        /// <summary>
        /// Saves an Excel workbook as a PDF file.
        /// </summary>
        public static void SaveAsPdf(this ExcelDocument document, string path, ExcelPdfSaveOptions? options = null) {
            document.ToPdfDocument(options).Save(path);
        }

        /// <summary>
        /// Writes an Excel workbook as PDF to a stream.
        /// </summary>
        public static void SaveAsPdf(this ExcelDocument document, Stream stream, ExcelPdfSaveOptions? options = null) {
            document.ToPdfDocument(options).Save(stream);
        }

        private static PdfCore.PdfOptions CreatePdfOptions(ExcelPdfSaveOptions options) {
            var pdfOptions = new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true
            };

            if (options.PageSize.HasValue) {
                pdfOptions.PageSize = options.PageSize.Value;
            }

            if (options.Margins.HasValue) {
                pdfOptions.Margins = options.Margins.Value;
            }

            return pdfOptions;
        }

        private static IReadOnlyList<WorksheetPdfExportPlan> BuildWorksheetExportPlans(ExcelDocument document, ExcelDocumentReader reader, IReadOnlyList<string> sheetNames, ExcelPdfSaveOptions options, bool hasExplicitSheetSelection) {
            var plans = new List<WorksheetPdfExportPlan>();
            for (int i = 0; i < sheetNames.Count; i++) {
                string sheetName = sheetNames[i];
                ExcelSheet? workbookSheet = GetWorkbookSheet(document, sheetName);
                if (ShouldSkipWorkbookSheet(workbookSheet, options, hasExplicitSheetSelection)) {
                    continue;
                }

                ExcelSheetReader sheet = reader.GetSheet(sheetName);
                ExcelSheetPageSetup? pageSetup = options.UseWorksheetPageSetup ? workbookSheet?.GetPageSetup() : null;
                ExcelSheet.HeaderFooterSnapshot? headerFooter = (options.UseWorksheetHeadersAndFooters || options.UseWorksheetHeaderFooterImages) ? workbookSheet?.GetHeaderFooter() : null;
                string exportRange = GetExportRange(sheet, workbookSheet, options);
                SheetExportData exportData = ReadSheetExportData(sheet, workbookSheet, exportRange, options);
                IReadOnlyList<int> manualRowBreaks = options.UseWorksheetPageBreaks && workbookSheet != null
                    ? workbookSheet.GetManualRowPageBreaks()
                    : Array.Empty<int>();
                IReadOnlyList<int> manualColumnBreaks = options.UseWorksheetPageBreaks && workbookSheet != null
                    ? workbookSheet.GetManualColumnPageBreaks()
                    : Array.Empty<int>();
                object?[,] values = exportData.Values;
                int rows = values.GetLength(0);
                int columns = values.GetLength(1);
                bool hasTable = rows > 0 && columns > 0;
                int exportedRows = options.MaxRowsPerSheet.HasValue
                    ? Math.Min(rows, options.MaxRowsPerSheet.Value)
                    : rows;
                if (options.MaxRowsPerSheet.HasValue && rows > exportedRows) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetRows",
                        $"Worksheet export was truncated from {rows.ToString(CultureInfo.InvariantCulture)} to {exportedRows.ToString(CultureInfo.InvariantCulture)} rows because MaxRowsPerSheet is set.");
                }

                ISet<string>? exportedCellReferences = CreateExportedCellReferenceSet(exportData.CellReferences, exportedRows);
                bool filterMediaToExportedCells = HasWorksheetPrintArea(workbookSheet, options) || options.MaxRowsPerSheet.HasValue;
                IReadOnlyList<WorksheetImageExportData> images = FilterImagesByExportedCells(ReadWorksheetImages(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                IReadOnlyList<WorksheetChartExportData> charts = FilterChartsByExportedCells(ReadWorksheetCharts(workbookSheet, options, sheetName), exportedCellReferences, filterMediaToExportedCells);
                if (!hasTable && images.Count == 0 && charts.Count == 0) {
                    continue;
                }

                plans.Add(new WorksheetPdfExportPlan(
                    sheetName,
                    pageSetup,
                    headerFooter,
                    exportData,
                    images,
                    charts,
                    hasTable,
                    exportedRows,
                    manualRowBreaks,
                    manualColumnBreaks,
                    CreateSheetBookmarkName(sheetName, plans.Count + 1)));
            }

            return plans;
        }

        private static IReadOnlyDictionary<string, string> BuildSheetDestinationMap(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var destinations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                destinations[plan.SheetName] = plan.BookmarkName;
            }

            return destinations;
        }

        private static IReadOnlyDictionary<string, string> BuildCellDestinationMap(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var targetCells = CollectInternalHyperlinkTargetCells(exportPlans);
            if (targetCells.Count == 0) {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            var destinations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                string?[,]? references = plan.ExportData.CellReferences;
                if (references == null) {
                    continue;
                }

                int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
                int columns = references.GetLength(1);
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        string? cellReference = references[row, column];
                        if (string.IsNullOrWhiteSpace(cellReference)) {
                            continue;
                        }

                        string key = CreateCellDestinationKey(plan.SheetName, cellReference!);
                        if (targetCells.Contains(key)) {
                            destinations[key] = CreateCellBookmarkName(plan.BookmarkName, cellReference!);
                        }
                    }
                }
            }

            return destinations;
        }

        private static HashSet<string> CollectInternalHyperlinkTargetCells(IReadOnlyList<WorksheetPdfExportPlan> exportPlans) {
            var targetCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (WorksheetPdfExportPlan plan in exportPlans) {
                ExcelHyperlinkSnapshot?[,]? hyperlinks = plan.ExportData.Hyperlinks;
                if (hyperlinks == null) {
                    continue;
                }

                int rows = Math.Min(plan.ExportedRows, hyperlinks.GetLength(0));
                int columns = hyperlinks.GetLength(1);
                for (int row = 0; row < rows; row++) {
                    for (int column = 0; column < columns; column++) {
                        ExcelHyperlinkSnapshot? hyperlink = hyperlinks[row, column];
                        if (hyperlink == null || hyperlink.IsExternal) {
                            continue;
                        }

                        if (TryParseInternalTarget(hyperlink.Target, plan.SheetName, out string? sheetName, out string? cellReference)) {
                            targetCells.Add(CreateCellDestinationKey(sheetName!, cellReference!));
                        }
                    }
                }
            }

            return targetCells;
        }

        private static string CreateSheetBookmarkName(string sheetName, int index) {
            var builder = new System.Text.StringBuilder();
            foreach (char value in sheetName) {
                if (char.IsLetterOrDigit(value)) {
                    builder.Append(char.ToLowerInvariant(value));
                } else if (builder.Length > 0 && builder[builder.Length - 1] != '-') {
                    builder.Append('-');
                }
            }

            string token = builder.ToString().Trim('-');
            if (token.Length > 40) {
                token = token.Substring(0, 40).Trim('-');
            }

            return token.Length == 0
                ? "excel-sheet-" + index.ToString(CultureInfo.InvariantCulture)
                : "excel-sheet-" + index.ToString(CultureInfo.InvariantCulture) + "-" + token;
        }

        private static string CreateCellBookmarkName(string sheetBookmarkName, string cellReference) {
            return sheetBookmarkName + "-" + cellReference.Replace("$", string.Empty).ToLowerInvariant();
        }

        private static string CreateCellDestinationKey(string sheetName, string cellReference) {
            return sheetName.Trim() + "!" + cellReference.Replace("$", string.Empty).ToUpperInvariant();
        }

        private static void ApplyWorksheetHeaderFooter(PdfCore.PdfPageCompose page, ExcelSheet.HeaderFooterSnapshot? headerFooter, string sheetName, string? workbookPath, ExcelPdfSaveOptions options) {
            if (headerFooter == null) {
                return;
            }

            WarnUnsupportedHeaderFooterImages(headerFooter, sheetName, options);

            HeaderFooterZones? headerZones = options.UseWorksheetHeadersAndFooters ? ConvertHeaderFooterZones(headerFooter.HeaderLeft, headerFooter.HeaderCenter, headerFooter.HeaderRight, sheetName, workbookPath, options, "header") : null;
            var firstHeaderZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage
                ? ConvertHeaderFooterZones(headerFooter.FirstHeaderLeft, headerFooter.FirstHeaderCenter, headerFooter.FirstHeaderRight, sheetName, workbookPath, options, "first-page header")
                : null;
            var evenHeaderZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven
                ? ConvertHeaderFooterZones(headerFooter.EvenHeaderLeft, headerFooter.EvenHeaderCenter, headerFooter.EvenHeaderRight, sheetName, workbookPath, options, "even-page header")
                : null;
            bool hasFirstHeaderVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage;
            bool hasEvenHeaderVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven;
            if (HasAnyText(headerZones) || hasFirstHeaderVariant || hasEvenHeaderVariant || HasAnyHeaderImage(headerFooter, options)) {
                page.Header(header => {
                    ApplyHeaderFooterStyle(header, ResolveSharedHeaderFooterStyle(new[] { headerZones, firstHeaderZones, evenHeaderZones }, sheetName, options, "header"));

                    if (HasAnyText(headerZones)) {
                        header.Zones(headerZones!.Left, headerZones.Center, headerZones.Right);
                    }

                    if (HasAnyText(firstHeaderZones)) {
                        header.FirstPageZones(firstHeaderZones!.Left, firstHeaderZones.Center, firstHeaderZones.Right);
                    } else if (hasFirstHeaderVariant) {
                        header.FirstPageText(string.Empty);
                    }

                    if (HasAnyText(evenHeaderZones)) {
                        header.EvenPagesZones(evenHeaderZones!.Left, evenHeaderZones.Center, evenHeaderZones.Right);
                    } else if (hasEvenHeaderVariant) {
                        header.EvenPagesText(string.Empty);
                    }

                    AddHeaderImage(header, headerFooter.HeaderLeftImage, options, PdfCore.PdfAlign.Left);
                    AddHeaderImage(header, headerFooter.HeaderCenterImage, options, PdfCore.PdfAlign.Center);
                    AddHeaderImage(header, headerFooter.HeaderRightImage, options, PdfCore.PdfAlign.Right);
                });
            }

            HeaderFooterZones? footerZones = options.UseWorksheetHeadersAndFooters ? ConvertHeaderFooterZones(headerFooter.FooterLeft, headerFooter.FooterCenter, headerFooter.FooterRight, sheetName, workbookPath, options, "footer") : null;
            var firstFooterZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage
                ? ConvertHeaderFooterZones(headerFooter.FirstFooterLeft, headerFooter.FirstFooterCenter, headerFooter.FirstFooterRight, sheetName, workbookPath, options, "first-page footer")
                : null;
            var evenFooterZones = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven
                ? ConvertHeaderFooterZones(headerFooter.EvenFooterLeft, headerFooter.EvenFooterCenter, headerFooter.EvenFooterRight, sheetName, workbookPath, options, "even-page footer")
                : null;
            bool hasFirstFooterVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentFirstPage;
            bool hasEvenFooterVariant = options.UseWorksheetHeadersAndFooters && headerFooter.DifferentOddEven;
            if (HasAnyText(footerZones) || hasFirstFooterVariant || hasEvenFooterVariant || HasAnyFooterImage(headerFooter, options)) {
                page.Footer(footer => {
                    ApplyHeaderFooterStyle(footer, ResolveSharedHeaderFooterStyle(new[] { footerZones, firstFooterZones, evenFooterZones }, sheetName, options, "footer"));

                    if (HasAnyText(footerZones)) {
                        footer.Zones(footerZones!.Left, footerZones.Center, footerZones.Right);
                    }

                    if (HasAnyText(firstFooterZones)) {
                        footer.FirstPageZones(firstFooterZones!.Left, firstFooterZones.Center, firstFooterZones.Right);
                    } else if (hasFirstFooterVariant) {
                        footer.FirstPageText(string.Empty);
                    }

                    if (HasAnyText(evenFooterZones)) {
                        footer.EvenPagesZones(evenFooterZones!.Left, evenFooterZones.Center, evenFooterZones.Right);
                    } else if (hasEvenFooterVariant) {
                        footer.EvenPagesText(string.Empty);
                    }

                    AddFooterImage(footer, headerFooter.FooterLeftImage, options, PdfCore.PdfAlign.Left);
                    AddFooterImage(footer, headerFooter.FooterCenterImage, options, PdfCore.PdfAlign.Center);
                    AddFooterImage(footer, headerFooter.FooterRightImage, options, PdfCore.PdfAlign.Right);
                });
            }
        }

        private static bool HasAnyHeaderImage(ExcelSheet.HeaderFooterSnapshot headerFooter, ExcelPdfSaveOptions options) {
            return options.UseWorksheetHeaderFooterImages
                   && (IsPdfSupportedImage(headerFooter.HeaderLeftImage)
                       || IsPdfSupportedImage(headerFooter.HeaderCenterImage)
                       || IsPdfSupportedImage(headerFooter.HeaderRightImage));
        }

        private static bool HasAnyFooterImage(ExcelSheet.HeaderFooterSnapshot headerFooter, ExcelPdfSaveOptions options) {
            return options.UseWorksheetHeaderFooterImages
                   && (IsPdfSupportedImage(headerFooter.FooterLeftImage)
                       || IsPdfSupportedImage(headerFooter.FooterCenterImage)
                       || IsPdfSupportedImage(headerFooter.FooterRightImage));
        }

        private static void AddHeaderImage(PdfCore.PdfHeaderCompose header, ExcelSheet.HeaderFooterImageSnapshot? image, ExcelPdfSaveOptions options, PdfCore.PdfAlign align) {
            if (options.UseWorksheetHeaderFooterImages && IsPdfSupportedImage(image)) {
                header.Image(image!.Bytes, image.WidthPoints, image.HeightPoints, align);
            }
        }

        private static void AddFooterImage(PdfCore.PdfFooterCompose footer, ExcelSheet.HeaderFooterImageSnapshot? image, ExcelPdfSaveOptions options, PdfCore.PdfAlign align) {
            if (options.UseWorksheetHeaderFooterImages && IsPdfSupportedImage(image)) {
                footer.Image(image!.Bytes, image.WidthPoints, image.HeightPoints, align);
            }
        }

        private static bool IsPdfSupportedImage(ExcelSheet.HeaderFooterImageSnapshot? image) {
            return image != null
                   && image.Bytes.Length > 0
                   && image.WidthPoints > 0D
                   && image.HeightPoints > 0D
                   && IsPdfSupportedImageContentType(image.ContentType)
                   && TryValidatePdfImageBytes(image.Bytes, image.ContentType, out _);
        }

        private static bool HasAnyText(params string?[] values) {
            foreach (string? value in values) {
                if (!string.IsNullOrWhiteSpace(value)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasAnyText(HeaderFooterZones? zones) {
            if (zones == null) {
                return false;
            }

            return HasAnyText(zones.Left, zones.Center, zones.Right);
        }

        private static void WarnUnsupportedHeaderFooterImages(ExcelSheet.HeaderFooterSnapshot headerFooter, string sheetName, ExcelPdfSaveOptions options) {
            if (!options.UseWorksheetHeaderFooterImages) {
                return;
            }

            WarnUnsupportedHeaderFooterImage(headerFooter.HeaderLeftImage, sheetName, "header left", options);
            WarnUnsupportedHeaderFooterImage(headerFooter.HeaderCenterImage, sheetName, "header center", options);
            WarnUnsupportedHeaderFooterImage(headerFooter.HeaderRightImage, sheetName, "header right", options);
            WarnUnsupportedHeaderFooterImage(headerFooter.FooterLeftImage, sheetName, "footer left", options);
            WarnUnsupportedHeaderFooterImage(headerFooter.FooterCenterImage, sheetName, "footer center", options);
            WarnUnsupportedHeaderFooterImage(headerFooter.FooterRightImage, sheetName, "footer right", options);
        }

        private static void WarnUnsupportedHeaderFooterImage(ExcelSheet.HeaderFooterImageSnapshot? image, string sheetName, string location, ExcelPdfSaveOptions options) {
            if (image == null || IsPdfSupportedImage(image)) {
                return;
            }

            AddWarning(
                options,
                sheetName,
                "WorksheetHeaderFooterImage",
                $"The {location} image was not exported because it is not a supported PDF image payload. ContentType='{image.ContentType}', WidthPoints={image.WidthPoints.ToString(CultureInfo.InvariantCulture)}, HeightPoints={image.HeightPoints.ToString(CultureInfo.InvariantCulture)}, Bytes={image.Bytes.Length.ToString(CultureInfo.InvariantCulture)}.");
        }

        private static HeaderFooterZones ConvertHeaderFooterZones(string? left, string? center, string? right, string sheetName, string? workbookPath, ExcelPdfSaveOptions options, string scope) {
            HeaderFooterZone leftZone = ConvertHeaderFooterText(left, sheetName, workbookPath, options, scope, "left");
            HeaderFooterZone centerZone = ConvertHeaderFooterText(center, sheetName, workbookPath, options, scope, "center");
            HeaderFooterZone rightZone = ConvertHeaderFooterText(right, sheetName, workbookPath, options, scope, "right");
            return new HeaderFooterZones(
                leftZone.Text,
                centerZone.Text,
                rightZone.Text,
                ResolveSharedHeaderFooterZoneStyle(new[] { leftZone, centerZone, rightZone }, sheetName, options, scope));
        }

        private static HeaderFooterZone ConvertHeaderFooterText(string? text, string sheetName, string? workbookPath, ExcelPdfSaveOptions options, string scope, string zone) {
            if (string.IsNullOrWhiteSpace(text)) {
                return HeaderFooterZone.Empty;
            }

            var builder = new System.Text.StringBuilder(text!.Length);
            var style = new HeaderFooterLineStyle();
            bool unsupportedFormatting = false;
            bool canApplyLineStyle = true;
            bool hasVisibleContent = false;
            DateTime? headerFooterDateTime = null;
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch != '&' || i + 1 >= text.Length) {
                    builder.Append(ch);
                    if (!char.IsWhiteSpace(ch)) {
                        hasVisibleContent = true;
                    }

                    continue;
                }

                char token = text[++i];
                switch (token) {
                    case '&':
                        if (i + 1 < text.Length && char.IsDigit(text[i + 1])) {
                            if (TryReadHeaderFooterFontSize(text, i + 1, out double fontSize, out int fontSizeEndIndex)) {
                                ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.FontSize = fontSize);
                                i = fontSizeEndIndex;
                            } else {
                                unsupportedFormatting = true;
                            }
                        } else {
                            builder.Append('&');
                            hasVisibleContent = true;
                        }

                        break;
                    case 'P':
                        builder.Append("{page}");
                        hasVisibleContent = true;
                        break;
                    case 'N':
                        builder.Append("{pages}");
                        hasVisibleContent = true;
                        break;
                    case 'A':
                        builder.Append(NormalizeHeaderFooterFieldText(sheetName));
                        hasVisibleContent = true;
                        break;
                    case 'D':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDateTime(options, ref headerFooterDateTime).ToString("d", CultureInfo.CurrentCulture)));
                        hasVisibleContent = true;
                        break;
                    case 'T':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDateTime(options, ref headerFooterDateTime).ToString("t", CultureInfo.CurrentCulture)));
                        hasVisibleContent = true;
                        break;
                    case 'F':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterFileName(workbookPath)));
                        hasVisibleContent = true;
                        break;
                    case 'Z':
                        builder.Append(NormalizeHeaderFooterFieldText(GetHeaderFooterDirectory(workbookPath)));
                        hasVisibleContent = true;
                        break;
                    case 'G':
                        break;
                    case 'B':
                        ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Bold = !s.Bold);
                        break;
                    case 'I':
                        ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Italic = !s.Italic);
                        break;
                    case 'U':
                    case 'S':
                        unsupportedFormatting = true;
                        if (hasVisibleContent) {
                            canApplyLineStyle = false;
                        }

                        break;
                    case 'K':
                        if (TryReadHeaderFooterColor(text, i + 1, out PdfCore.PdfColor color, out int colorEndIndex)) {
                            ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.Color = color);
                            i = colorEndIndex;
                        } else {
                            unsupportedFormatting = true;
                            i = SkipExcelHeaderFooterColorToken(text, i);
                        }

                        break;
                    case '"':
                        if (TryReadHeaderFooterQuotedToken(text, i, out string quotedToken, out int quotedEndIndex) &&
                            TryApplyHeaderFooterFontToken(style, quotedToken)) {
                            ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, _ => { });
                            i = quotedEndIndex;
                        } else {
                            unsupportedFormatting = true;
                            i = SkipExcelHeaderFooterQuotedToken(text, i);
                        }

                        break;
                    default:
                        if (char.IsDigit(token)) {
                            if (TryReadHeaderFooterFontSize(text, i, out double fontSize, out int fontSizeEndIndex)) {
                                ApplyHeaderFooterStyleToken(style, hasVisibleContent, ref unsupportedFormatting, ref canApplyLineStyle, s => s.FontSize = fontSize);
                                i = fontSizeEndIndex;
                            } else {
                                unsupportedFormatting = true;
                            }
                        } else {
                            builder.Append(token);
                            hasVisibleContent = true;
                        }
                        break;
                }
            }

            if (unsupportedFormatting) {
                AddWarning(
                    options,
                    sheetName,
                    "WorksheetHeaderFooterFormatting",
                    $"Excel header/footer formatting in the {scope} {zone} zone was simplified. Text, page tokens, total-page tokens, sheet-name, date/time, and workbook file fields are preserved, but rich formatting is not exported yet.");
            }

            string result = builder.ToString().Trim();
            return result.Length == 0
                ? HeaderFooterZone.Empty
                : new HeaderFooterZone(result, canApplyLineStyle && style.HasAnyStyle ? style : null);
        }

        private static void ApplyHeaderFooterStyleToken(HeaderFooterLineStyle style, bool hasVisibleContent, ref bool unsupportedFormatting, ref bool canApplyLineStyle, Action<HeaderFooterLineStyle> apply) {
            if (hasVisibleContent) {
                unsupportedFormatting = true;
                canApplyLineStyle = false;
                return;
            }

            apply(style);
        }

        private static HeaderFooterLineStyle? ResolveSharedHeaderFooterZoneStyle(HeaderFooterZone[] zones, string sheetName, ExcelPdfSaveOptions options, string scope) {
            HeaderFooterLineStyle? shared = null;
            bool hasStyle = false;
            bool hasUnstyledText = false;
            foreach (HeaderFooterZone zone in zones) {
                if (string.IsNullOrWhiteSpace(zone.Text)) {
                    continue;
                }

                if (zone.Style == null) {
                    hasUnstyledText = true;
                    continue;
                }

                if (!hasStyle) {
                    shared = zone.Style;
                    hasStyle = true;
                    continue;
                }

                if (!HeaderFooterLineStyle.Equals(shared, zone.Style)) {
                    AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                    return null;
                }
            }

            if (hasStyle && hasUnstyledText) {
                AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                return null;
            }

            return shared;
        }

        private static HeaderFooterLineStyle? ResolveSharedHeaderFooterStyle(HeaderFooterZones?[] zoneSets, string sheetName, ExcelPdfSaveOptions options, string scope) {
            HeaderFooterLineStyle? shared = null;
            bool hasStyle = false;
            bool hasUnstyledText = false;
            foreach (HeaderFooterZones? zones in zoneSets) {
                if (!HasAnyText(zones)) {
                    continue;
                }

                if (zones!.Style == null) {
                    hasUnstyledText = true;
                    continue;
                }

                if (!hasStyle) {
                    shared = zones.Style;
                    hasStyle = true;
                    continue;
                }

                if (!HeaderFooterLineStyle.Equals(shared, zones.Style)) {
                    AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                    return null;
                }
            }

            if (hasStyle && hasUnstyledText) {
                AddMixedHeaderFooterFormattingWarning(options, sheetName, scope);
                return null;
            }

            return shared;
        }

        private static void ApplyHeaderFooterStyle(PdfCore.PdfHeaderCompose header, HeaderFooterLineStyle? style) {
            if (style == null) {
                return;
            }

            if (style.FontSize.HasValue) {
                header.FontSize(style.FontSize.Value);
            }

            if (style.Color.HasValue) {
                header.Color(style.Color.Value);
            }

            if (style.Font.HasValue) {
                header.Font(style.Font.Value);
            }
        }

        private static void ApplyHeaderFooterStyle(PdfCore.PdfFooterCompose footer, HeaderFooterLineStyle? style) {
            if (style == null) {
                return;
            }

            if (style.FontSize.HasValue) {
                footer.FontSize(style.FontSize.Value);
            }

            if (style.Color.HasValue) {
                footer.Color(style.Color.Value);
            }

            if (style.Font.HasValue) {
                footer.Font(style.Font.Value);
            }
        }

        private static void AddMixedHeaderFooterFormattingWarning(ExcelPdfSaveOptions options, string sheetName, string scope) {
            AddWarning(
                options,
                sheetName,
                "WorksheetHeaderFooterFormatting",
                $"Excel header/footer formatting in the {scope} uses mixed or partial styles that cannot be represented as one PDF header/footer line style yet. Text is preserved, but rich formatting is simplified.");
        }

        private static DateTime GetHeaderFooterDateTime(ExcelPdfSaveOptions options, ref DateTime? dateTime) {
            if (!dateTime.HasValue) {
                dateTime = options.HeaderFooterDateTimeProvider != null
                    ? options.HeaderFooterDateTimeProvider()
                    : DateTime.Now;
            }

            return dateTime.Value;
        }

        private static string GetHeaderFooterFileName(string? workbookPath) {
            if (string.IsNullOrWhiteSpace(workbookPath)) {
                return "Workbook";
            }

            string fileName = Path.GetFileName(workbookPath!);
            return string.IsNullOrWhiteSpace(fileName) ? "Workbook" : fileName;
        }

        private static string GetHeaderFooterDirectory(string? workbookPath) {
            if (string.IsNullOrWhiteSpace(workbookPath)) {
                return string.Empty;
            }

            string? directory = Path.GetDirectoryName(Path.GetFullPath(workbookPath!));
            return directory ?? string.Empty;
        }

        private static string NormalizeHeaderFooterFieldText(string text) =>
            text.Replace('\u00A0', ' ').Replace('\u202F', ' ');

        private static bool TryReadHeaderFooterFontSize(string text, int startIndex, out double fontSize, out int endIndex) {
            fontSize = 0D;
            endIndex = startIndex;
            int index = startIndex;
            while (index < text.Length && char.IsDigit(text[index])) {
                index++;
            }

            if (index == startIndex) {
                return false;
            }

            string token = text.Substring(startIndex, index - startIndex);
            endIndex = index - 1;
            return double.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out fontSize) && fontSize > 0D;
        }

        private static bool TryReadHeaderFooterColor(string text, int startIndex, out PdfCore.PdfColor color, out int endIndex) {
            color = default;
            endIndex = startIndex;
            if (startIndex + 6 > text.Length) {
                return false;
            }

            for (int i = 0; i < 6; i++) {
                if (!IsHexDigit(text[startIndex + i])) {
                    return false;
                }
            }

            byte red = byte.Parse(text.Substring(startIndex, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte green = byte.Parse(text.Substring(startIndex + 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte blue = byte.Parse(text.Substring(startIndex + 4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            color = PdfCore.PdfColor.FromRgb(red, green, blue);
            endIndex = startIndex + 5;
            return true;
        }

        private static bool TryReadHeaderFooterQuotedToken(string text, int quoteIndex, out string value, out int endIndex) {
            value = string.Empty;
            endIndex = quoteIndex;
            int closingIndex = quoteIndex + 1;
            while (closingIndex < text.Length && text[closingIndex] != '"') {
                closingIndex++;
            }

            if (closingIndex >= text.Length) {
                return false;
            }

            value = text.Substring(quoteIndex + 1, closingIndex - quoteIndex - 1);
            endIndex = closingIndex;
            return true;
        }

        private static bool TryApplyHeaderFooterFontToken(HeaderFooterLineStyle style, string token) {
            if (string.IsNullOrWhiteSpace(token)) {
                return false;
            }

            string[] parts = token.Split(new[] { ',' }, 2);
            if (!TryMapHeaderFooterFontFamily(parts[0], out HeaderFooterFontFamily family)) {
                return false;
            }

            style.FontFamily = family;
            if (parts.Length > 1) {
                string fontStyle = parts[1];
                style.Bold = fontStyle.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0;
                style.Italic = fontStyle.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                               fontStyle.IndexOf("oblique", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return true;
        }

        private static bool TryMapHeaderFooterFontFamily(string value, out HeaderFooterFontFamily family) {
            string normalized = NormalizeHeaderFooterFontName(value);
            if (normalized == "arial" || normalized == "calibri" || normalized == "helvetica") {
                family = HeaderFooterFontFamily.Helvetica;
                return true;
            }

            if (normalized == "times" || normalized == "timesnewroman") {
                family = HeaderFooterFontFamily.Times;
                return true;
            }

            if (normalized == "courier" || normalized == "couriernew") {
                family = HeaderFooterFontFamily.Courier;
                return true;
            }

            family = HeaderFooterFontFamily.Helvetica;
            return false;
        }

        private static string NormalizeHeaderFooterFontName(string value) {
            var builder = new System.Text.StringBuilder(value.Length);
            foreach (char ch in value) {
                if (char.IsLetterOrDigit(ch)) {
                    builder.Append(char.ToLowerInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private static int SkipExcelHeaderFooterColorToken(string text, int index) {
            int skipped = 0;
            while (index + 1 < text.Length && skipped < 6 && IsHexDigit(text[index + 1])) {
                index++;
                skipped++;
            }

            return index;
        }

        private static int SkipExcelHeaderFooterQuotedToken(string text, int index) {
            while (index + 1 < text.Length) {
                index++;
                if (text[index] == '"') {
                    break;
                }
            }

            return index;
        }

        private static int SkipExcelHeaderFooterFontSizeToken(string text, int index) {
            while (index + 1 < text.Length && char.IsDigit(text[index + 1])) {
                index++;
            }

            return index;
        }

        private static bool IsHexDigit(char value) {
            return (value >= '0' && value <= '9') ||
                   (value >= 'a' && value <= 'f') ||
                   (value >= 'A' && value <= 'F');
        }

        private static void ApplyWorksheetPageSetup(PdfCore.PdfPageCompose page, ExcelSheetPageSetup? pageSetup, ExcelPdfSaveOptions options) {
            PdfCore.PageSize? pageSize = options.PageSize;
            if (!options.PageSize.HasValue && pageSetup?.Orientation == ExcelPageOrientation.Landscape) {
                pageSize = (pageSize ?? PdfCore.PageSizes.Letter).Landscape();
            } else if (!options.PageSize.HasValue && pageSetup?.Orientation == ExcelPageOrientation.Portrait) {
                pageSize = (pageSize ?? PdfCore.PageSizes.Letter).Portrait();
            }

            if (pageSize.HasValue) {
                page.Size(pageSize.Value);
            }

            if (options.Margins.HasValue) {
                page.Margin(options.Margins.Value);
            } else if (pageSetup?.Margins != null) {
                page.Margin(ToPdfMargins(pageSetup.Margins));
            }
        }

        private static PdfCore.PageMargins ToPdfMargins(ExcelSheetPageMargins margins) {
            return PdfCore.PageMargins.FromInches(margins.Left, margins.Top, margins.Right, margins.Bottom);
        }

        private static IReadOnlyList<string> GetSheetNames(ExcelDocumentReader reader, ExcelPdfSaveOptions options) {
            IReadOnlyList<string> requestedNames = options.SheetNames ?? Array.Empty<string>();
            if (requestedNames.Count == 0) {
                return reader.GetSheetNames();
            }

            var names = new List<string>(requestedNames.Count);
            foreach (string name in requestedNames) {
                if (string.IsNullOrWhiteSpace(name)) {
                    throw new ArgumentException("Sheet names cannot contain null, empty, or whitespace values.", nameof(options));
                }

                reader.GetSheet(name);
                names.Add(name);
            }

            return names;
        }

        private static bool HasExplicitSheetSelection(ExcelPdfSaveOptions options) {
            return options.SheetNames != null && options.SheetNames.Count > 0;
        }

        private static bool ShouldSkipWorkbookSheet(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, bool hasExplicitSheetSelection) {
            return !hasExplicitSheetSelection
                   && options.RespectWorkbookSheetVisibility
                   && workbookSheet?.Hidden == true;
        }

        private static IReadOnlyList<WorksheetImageExportData> ReadWorksheetImages(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, string sheetName) {
            if (!options.UseWorksheetImages || workbookSheet == null) {
                return Array.Empty<WorksheetImageExportData>();
            }

            var images = new List<WorksheetImageExportData>();
            foreach (ExcelImage image in workbookSheet.Images.OrderBy(image => image.RowIndex).ThenBy(image => image.ColumnIndex)) {
                if (!IsPdfSupportedImageContentType(image.ContentType)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because content type '{image.ContentType}' is not supported by the first-party PDF image writer.");
                    continue;
                }

                byte[] bytes = image.GetBytes();
                if (bytes.Length == 0 || image.WidthPixels <= 0 || image.HeightPixels <= 0) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because it has empty image bytes or non-positive dimensions.");
                    continue;
                }

                if (!TryValidatePdfImageBytes(bytes, image.ContentType, out string? unsupportedReason)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetImage",
                        $"Worksheet image anchored at {A1.CellReference(image.RowIndex, image.ColumnIndex)} was not exported because the first-party PDF image writer cannot export the image bytes. {unsupportedReason}");
                    continue;
                }

                images.Add(new WorksheetImageExportData(bytes, PixelsToPoints(image.WidthPixels), PixelsToPoints(image.HeightPixels), A1.CellReference(image.RowIndex, image.ColumnIndex)));
            }

            return images;
        }

        private static IReadOnlyList<WorksheetChartExportData> ReadWorksheetCharts(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options, string sheetName) {
            if (!options.UseWorksheetCharts || workbookSheet == null) {
                return Array.Empty<WorksheetChartExportData>();
            }

            var charts = new List<WorksheetChartExportData>();
            foreach (ExcelChart chart in workbookSheet.Charts) {
                if (!chart.TryGetSnapshot(out ExcelChartSnapshot snapshot)) {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{chart.Name}' was not exported because its chart data could not be read into a first-party PDF snapshot.");
                    continue;
                }

                if (IsSupportedChartSnapshot(snapshot)) {
                    charts.Add(new WorksheetChartExportData(snapshot));
                } else {
                    AddWarning(
                        options,
                        sheetName,
                        "WorksheetChart",
                        $"Worksheet chart '{GetChartDisplayName(snapshot)}' was not exported because chart type '{snapshot.ChartType}' is not supported by the first-party PDF chart snapshot renderer yet.");
                }
            }

            return charts
                .OrderBy(chart => chart.Snapshot.RowIndex)
                .ThenBy(chart => chart.Snapshot.ColumnIndex)
                .ToList();
        }

        private static bool IsSupportedChartSnapshot(ExcelChartSnapshot snapshot) {
            return IsColumnChart(snapshot.ChartType)
                   || IsBarChart(snapshot.ChartType)
                   || IsLineChart(snapshot.ChartType)
                   || IsAreaChart(snapshot.ChartType)
                   || IsScatterChart(snapshot.ChartType)
                   || IsRadarChart(snapshot.ChartType)
                   || IsPieChart(snapshot.ChartType)
                   || IsDoughnutChart(snapshot.ChartType);
        }

        private static string GetChartDisplayName(ExcelChartSnapshot snapshot) {
            if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
                return snapshot.Title!;
            }

            return string.IsNullOrWhiteSpace(snapshot.Name) ? snapshot.ChartType.ToString() : snapshot.Name;
        }

        private static void AddWarning(ExcelPdfSaveOptions options, string sheetName, string feature, string message) {
            options.Warnings.Add(new ExcelPdfExportWarning(sheetName, feature, message));
        }

        private static bool IsPdfSupportedImageContentType(string contentType) {
            return string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase)
                   || string.Equals(contentType, "image/jpg", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryValidatePdfImageBytes(byte[] bytes, string contentType, out string? unsupportedReason) {
            unsupportedReason = null;
            if (!OfficeImageReader.TryIdentify(bytes, null, out OfficeImageInfo imageInfo)) {
                unsupportedReason = "Image bytes do not contain a supported image header.";
                return false;
            }

            if (IsPngContentType(contentType) && imageInfo.Format != OfficeImageFormat.Png) {
                unsupportedReason = $"Image bytes were declared as PNG but were detected as {imageInfo.Format}.";
                return false;
            }

            if (IsJpegContentType(contentType) && imageInfo.Format != OfficeImageFormat.Jpeg) {
                unsupportedReason = $"Image bytes were declared as JPEG but were detected as {imageInfo.Format}.";
                return false;
            }

            try {
                _ = new PdfCore.PdfTableCellImage(bytes, 1D, 1D);
                return true;
            } catch (ArgumentException ex) {
                unsupportedReason = ex.Message;
                return false;
            } catch (InvalidDataException ex) {
                unsupportedReason = ex.Message;
                return false;
            } catch (NotSupportedException ex) {
                unsupportedReason = ex.Message;
                return false;
            }
        }

        private static bool IsPngContentType(string contentType) =>
            string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase);

        private static bool IsJpegContentType(string contentType) =>
            string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase)
            || string.Equals(contentType, "image/jpg", StringComparison.OrdinalIgnoreCase);

        private static double PixelsToPoints(int pixels) {
            return pixels * 72D / 96D;
        }

        private enum HeaderFooterFontFamily {
            Helvetica,
            Times,
            Courier
        }

        private sealed class HeaderFooterZone {
            internal static readonly HeaderFooterZone Empty = new HeaderFooterZone(null, null);

            internal HeaderFooterZone(string? text, HeaderFooterLineStyle? style) {
                Text = text;
                Style = style;
            }

            internal string? Text { get; }

            internal HeaderFooterLineStyle? Style { get; }
        }

        private sealed class HeaderFooterZones {
            internal HeaderFooterZones(string? left, string? center, string? right, HeaderFooterLineStyle? style) {
                Left = left;
                Center = center;
                Right = right;
                Style = style;
            }

            internal string? Left { get; }

            internal string? Center { get; }

            internal string? Right { get; }

            internal HeaderFooterLineStyle? Style { get; }
        }

        private sealed class HeaderFooterLineStyle {
            internal double? FontSize { get; set; }

            internal PdfCore.PdfColor? Color { get; set; }

            internal HeaderFooterFontFamily? FontFamily { get; set; }

            internal bool Bold { get; set; }

            internal bool Italic { get; set; }

            internal PdfCore.PdfStandardFont? Font {
                get {
                    if (!FontFamily.HasValue && !Bold && !Italic) {
                        return null;
                    }

                    HeaderFooterFontFamily family = FontFamily ?? HeaderFooterFontFamily.Helvetica;
                    switch (family) {
                        case HeaderFooterFontFamily.Times:
                            if (Bold && Italic) return PdfCore.PdfStandardFont.TimesBoldItalic;
                            if (Bold) return PdfCore.PdfStandardFont.TimesBold;
                            if (Italic) return PdfCore.PdfStandardFont.TimesItalic;
                            return PdfCore.PdfStandardFont.TimesRoman;
                        case HeaderFooterFontFamily.Courier:
                            if (Bold && Italic) return PdfCore.PdfStandardFont.CourierBoldOblique;
                            if (Bold) return PdfCore.PdfStandardFont.CourierBold;
                            if (Italic) return PdfCore.PdfStandardFont.CourierOblique;
                            return PdfCore.PdfStandardFont.Courier;
                        default:
                            if (Bold && Italic) return PdfCore.PdfStandardFont.HelveticaBoldOblique;
                            if (Bold) return PdfCore.PdfStandardFont.HelveticaBold;
                            if (Italic) return PdfCore.PdfStandardFont.HelveticaOblique;
                            return PdfCore.PdfStandardFont.Helvetica;
                    }
                }
            }

            internal bool HasAnyStyle {
                get {
                    PdfCore.PdfStandardFont? font = Font;
                    return FontSize.HasValue || Color.HasValue || (font.HasValue && font.Value != PdfCore.PdfStandardFont.Helvetica);
                }
            }

            internal static bool Equals(HeaderFooterLineStyle? left, HeaderFooterLineStyle? right) {
                if (ReferenceEquals(left, right)) {
                    return true;
                }

                if (left == null || right == null) {
                    return false;
                }

                return Nullable.Equals(left.FontSize, right.FontSize)
                       && Nullable.Equals(left.Color, right.Color)
                       && Nullable.Equals(left.Font, right.Font);
            }
        }

        private static void AddWorksheetChart(PdfCore.PdfItemCompose item, WorksheetChartExportData chart) {
            ExcelChartSnapshot snapshot = chart.Snapshot;
            string title = string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title!;
            if (!string.IsNullOrWhiteSpace(title)) {
                item.H2(title, PdfCore.PdfAlign.Left, PdfCore.PdfColor.FromRgb(31, 78, 121));
            }

            item.Drawing(CreateChartDrawing(snapshot), PdfCore.PdfAlign.Left, spacingBefore: 2, spacingAfter: 6);
            item.Table(CreateChartLegendRows(snapshot), PdfCore.PdfAlign.Left, CreateChartLegendStyle(GetChartLegendColorCount(snapshot)));
        }

        private static OfficeDrawing CreateChartDrawing(ExcelChartSnapshot snapshot) {
            double width = Math.Min(420D, Math.Max(240D, PixelsToPoints(snapshot.WidthPixels)));
            double height = Math.Min(260D, Math.Max(150D, PixelsToPoints(snapshot.HeightPixels)));
            var drawing = new OfficeDrawing(width, height);

            AddShape(drawing, OfficeShape.Rectangle(width, height), 0, 0, OfficeColor.FromRgb(250, 252, 255), OfficeColor.FromRgb(183, 194, 207), 0.75);

            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                AddPieSeries(drawing, snapshot, width, height, IsDoughnutChart(snapshot.ChartType));
                return drawing;
            }

            if (IsRadarChart(snapshot.ChartType)) {
                AddRadarSeries(drawing, snapshot, width, height);
                return drawing;
            }

            double plotLeft = 36D;
            double plotTop = 18D;
            double plotRight = 12D;
            double plotBottom = 28D;
            double plotWidth = Math.Max(20D, width - plotLeft - plotRight);
            double plotHeight = Math.Max(20D, height - plotTop - plotBottom);
            double plotBottomY = plotTop + plotHeight;

            AddShape(drawing, OfficeShape.Line(0, 0, plotWidth, 0), plotLeft, plotBottomY, null, OfficeColor.FromRgb(80, 90, 100), 0.75);
            AddShape(drawing, OfficeShape.Line(0, 0, 0, plotHeight), plotLeft, plotTop, null, OfficeColor.FromRgb(80, 90, 100), 0.75);
            for (int i = 1; i <= 3; i++) {
                double y = plotTop + (plotHeight * i / 4D);
                AddShape(drawing, OfficeShape.Line(0, 0, plotWidth, 0), plotLeft, y, null, OfficeColor.FromRgb(226, 232, 240), 0.5);
            }

            if (IsAreaChart(snapshot.ChartType)) {
                AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else if (IsScatterChart(snapshot.ChartType)) {
                AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else if (IsLineChart(snapshot.ChartType)) {
                AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else {
                AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            }

            return drawing;
        }

        private static void AddBarSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            double slot = plotWidth / categories.Count;
            double groupWidth = slot * 0.68D;
            bool horizontal = IsBarChart(snapshot.ChartType);
            bool stacked = IsStackedBarOrColumnChart(snapshot.ChartType) || IsPercentStackedBarOrColumnChart(snapshot.ChartType);
            bool percentStacked = IsPercentStackedBarOrColumnChart(snapshot.ChartType);
            double barWidth = Math.Max(2D, stacked ? groupWidth : groupWidth / series.Count);
            (double min, double max) = percentStacked
                ? (0D, 1D)
                : stacked
                    ? GetStackedSeriesRange(series, categories.Count)
                    : GetFiniteSeriesRange(series);
            min = Math.Min(0D, min);
            max = Math.Max(0D, max);
            if (max <= min) {
                max = min + 1D;
            }

            for (int category = 0; category < categories.Count; category++) {
                double positiveBase = 0D;
                double negativeBase = 0D;
                double percentTotal = percentStacked ? GetPositiveCategoryTotal(series, category) : 0D;
                for (int s = 0; s < series.Count; s++) {
                    double value = GetSeriesValue(series[s], category);
                    if (value == 0D) {
                        continue;
                    }

                    double baseline = 0D;
                    double plottedValue = value;
                    if (stacked) {
                        if (percentStacked) {
                            plottedValue = percentTotal <= 0D ? 0D : Math.Max(0D, value) / percentTotal;
                        }

                        baseline = plottedValue >= 0D ? positiveBase : negativeBase;
                        if (plottedValue >= 0D) {
                            positiveBase += plottedValue;
                        } else {
                            negativeBase += plottedValue;
                        }
                    }

                    OfficeColor color = GetChartSeriesColor(s);
                    if (horizontal) {
                        double categoryHeight = plotHeight / categories.Count;
                        double rowHeight = Math.Max(2D, categoryHeight * 0.68D / (stacked ? 1D : series.Count));
                        double y = plotTop + (categoryHeight * category) + (categoryHeight * 0.16D) + (stacked ? 0D : rowHeight * s);
                        double x1 = ToPlotX(baseline, min, max, plotLeft, plotWidth);
                        double x2 = ToPlotX(stacked ? baseline + plottedValue : plottedValue, min, max, plotLeft, plotWidth);
                        double x = Math.Min(x1, x2);
                        double w = Math.Max(1D, Math.Abs(x2 - x1));
                        AddShape(drawing, OfficeShape.Rectangle(w, rowHeight), x, y, color, null, 0);
                    } else {
                        double x = plotLeft + (slot * category) + ((slot - groupWidth) / 2D) + (stacked ? 0D : barWidth * s);
                        double y1 = ToPlotY(baseline, min, max, plotTop, plotHeight);
                        double y2 = ToPlotY(stacked ? baseline + plottedValue : plottedValue, min, max, plotTop, plotHeight);
                        double y = Math.Min(y1, y2);
                        double h = Math.Max(1D, Math.Abs(y2 - y1));
                        AddShape(drawing, OfficeShape.Rectangle(barWidth * 0.88D, h), x, y, color, null, 0);
                    }
                }
            }
        }

        private static void AddAreaSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 2 || series.Count == 0) {
                return;
            }

            bool stacked = IsStackedAreaChart(snapshot.ChartType) || IsPercentStackedAreaChart(snapshot.ChartType);
            bool percentStacked = IsPercentStackedAreaChart(snapshot.ChartType);
            double max = percentStacked ? 1D : stacked ? GetPositiveStackedMax(series, categories.Count) : GetPositiveMax(series);
            double step = plotWidth / (categories.Count - 1);
            var cumulative = new double[categories.Count];

            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var topPoints = new List<OfficePoint>(categories.Count);
                var bottomPoints = new List<OfficePoint>(categories.Count);

                for (int i = 0; i < categories.Count; i++) {
                    double rawValue = Math.Max(0D, GetSeriesValue(series[s], i));
                    double baseline = stacked ? cumulative[i] : 0D;
                    double topValue = baseline + rawValue;

                    if (percentStacked) {
                        double total = GetPositiveCategoryTotal(series, i);
                        baseline = total <= 0D ? 0D : baseline / total;
                        topValue = total <= 0D ? 0D : topValue / total;
                    }

                    double x = plotLeft + step * i;
                    topPoints.Add(new OfficePoint(x, ToPlotY(topValue, max, plotTop, plotHeight)));
                    bottomPoints.Add(new OfficePoint(x, ToPlotY(baseline, max, plotTop, plotHeight)));
                }

                var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
                areaPoints.AddRange(topPoints);
                for (int i = bottomPoints.Count - 1; i >= 0; i--) {
                    areaPoints.Add(bottomPoints[i]);
                }

                AddPolygonShape(drawing, areaPoints, color, color, 0.5D, 0.32D);
                AddPointLine(drawing, topPoints, color, 1.4D);

                if (stacked) {
                    for (int i = 0; i < categories.Count; i++) {
                        cumulative[i] += Math.Max(0D, GetSeriesValue(series[s], i));
                    }
                }
            }
        }

        private static void AddLineSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 2 || series.Count == 0) {
                return;
            }

            (double min, double max) = GetFiniteSeriesRange(series);
            double step = plotWidth / (categories.Count - 1);
            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                for (int i = 1; i < categories.Count; i++) {
                    double x1 = plotLeft + step * (i - 1);
                    double y1 = ToPlotY(GetSeriesValue(series[s], i - 1), min, max, plotTop, plotHeight);
                    double x2 = plotLeft + step * i;
                    double y2 = ToPlotY(GetSeriesValue(series[s], i), min, max, plotTop, plotHeight);
                    double minX = Math.Min(x1, x2);
                    double minY = Math.Min(y1, y2);
                    AddShape(drawing, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, color, 1.75);
                }

                for (int i = 0; i < categories.Count; i++) {
                    double x = plotLeft + step * i - 2D;
                    double y = ToPlotY(GetSeriesValue(series[s], i), min, max, plotTop, plotHeight) - 2D;
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), x, y, OfficeColor.White, color, 1D);
                }
            }
        }

        private static void AddScatterSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            IReadOnlyList<double> xValues = GetScatterXValues(categories);
            (double minX, double maxX) = GetFiniteRange(xValues);
            (double minY, double maxY) = GetFiniteSeriesRange(series);
            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var points = new List<OfficePoint>(categories.Count);
                for (int i = 0; i < categories.Count; i++) {
                    double yValue = GetSeriesValue(series[s], i);
                    double x = ToPlotX(xValues[i], minX, maxX, plotLeft, plotWidth);
                    double y = ToPlotY(yValue, minY, maxY, plotTop, plotHeight);
                    points.Add(new OfficePoint(x, y));
                }

                AddPointLine(drawing, points, color, 1.25D);
                for (int i = 0; i < points.Count; i++) {
                    OfficePoint point = points[i];
                    AddShape(drawing, OfficeShape.Ellipse(5D, 5D), point.X - 2.5D, point.Y - 2.5D, OfficeColor.White, color, 1.25D);
                }
            }
        }

        private static void AddRadarSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double width, double height) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 3 || series.Count == 0) {
                return;
            }

            double centerX = width / 2D;
            double centerY = height / 2D;
            double radius = Math.Max(36D, Math.Min(width - 52D, height - 42D) / 2D);
            double max = GetPositiveMax(series);

            for (int ring = 1; ring <= 4; ring++) {
                double ringRadius = radius * ring / 4D;
                IReadOnlyList<OfficePoint> ringPoints = CreateRadarPoints(categories.Count, centerX, centerY, ringRadius);
                AddPolygonShape(drawing, ringPoints, null, OfficeColor.FromRgb(226, 232, 240), 0.5D);
            }

            IReadOnlyList<OfficePoint> outerPoints = CreateRadarPoints(categories.Count, centerX, centerY, radius);
            for (int i = 0; i < outerPoints.Count; i++) {
                OfficePoint point = outerPoints[i];
                double minX = Math.Min(centerX, point.X);
                double minY = Math.Min(centerY, point.Y);
                AddShape(
                    drawing,
                    OfficeShape.Line(centerX - minX, centerY - minY, point.X - minX, point.Y - minY),
                    minX,
                    minY,
                    null,
                    OfficeColor.FromRgb(203, 213, 225),
                    0.5D);
            }

            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var points = new List<OfficePoint>(categories.Count);
                for (int i = 0; i < categories.Count; i++) {
                    double value = Math.Max(0D, GetSeriesValue(series[s], i));
                    double pointRadius = radius * Math.Min(1D, value / max);
                    points.Add(CreateRadarPoint(i, categories.Count, centerX, centerY, pointRadius));
                }

                AddPolygonShape(drawing, points, color, color, 1D, 0.18D);
                for (int i = 0; i < points.Count; i++) {
                    OfficePoint point = points[i];
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), point.X - 2D, point.Y - 2D, OfficeColor.White, color, 1D);
                }
            }
        }

        private static void AddPieSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double width, double height, bool doughnut) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            ExcelChartSeries values = series[0];
            double total = 0D;
            for (int i = 0; i < categories.Count; i++) {
                double value = GetSeriesValue(values, i);
                if (!double.IsNaN(value) && !double.IsInfinity(value) && value > 0D) {
                    total += value;
                }
            }

            if (total <= 0D) {
                return;
            }

            double radius = Math.Max(36D, Math.Min(width - 48D, height - 36D) / 2D);
            double centerX = width / 2D;
            double centerY = height / 2D;
            double start = -Math.PI / 2D;
            for (int i = 0; i < categories.Count; i++) {
                double value = Math.Max(0D, GetSeriesValue(values, i));
                if (value <= 0D) {
                    continue;
                }

                double sweep = value / total * Math.PI * 2D;
                double end = start + sweep;
                var points = new List<OfficePoint> {
                    new OfficePoint(centerX, centerY)
                };
                int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
                for (int segment = 0; segment <= segments; segment++) {
                    double angle = start + (sweep * segment / segments);
                    points.Add(new OfficePoint(
                        centerX + Math.Cos(angle) * radius,
                        centerY + Math.Sin(angle) * radius));
                }

                AddPolygonShape(drawing, points, GetChartSeriesColor(i), OfficeColor.White, 0.5D);
                start = end;
            }

            if (doughnut) {
                double innerDiameter = radius * 1.02D;
                AddShape(
                    drawing,
                    OfficeShape.Ellipse(innerDiameter, innerDiameter),
                    centerX - innerDiameter / 2D,
                    centerY - innerDiameter / 2D,
                    OfficeColor.FromRgb(250, 252, 255),
                    null,
                    0D);
            }
        }

        private static string[][] CreateChartLegendRows(ExcelChartSnapshot snapshot) {
            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                return CreatePieChartLegendRows(snapshot);
            }

            var rows = new List<string[]> {
                new[] { "Series", "Values" }
            };

            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                rows.Add(new[] {
                    series.Name,
                    string.Join(", ", series.Values.Select(value => value.ToString("0.##", CultureInfo.InvariantCulture)))
                });
            }

            return rows.ToArray();
        }

        private static string[][] CreatePieChartLegendRows(ExcelChartSnapshot snapshot) {
            var rows = new List<string[]> {
                new[] { "Category", "Value" }
            };

            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            ExcelChartSeries? values = series.Count > 0 ? series[0] : null;
            for (int i = 0; i < categories.Count; i++) {
                string category = string.IsNullOrWhiteSpace(categories[i])
                    ? "Slice " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    : categories[i];
                rows.Add(new[] {
                    category,
                    values == null ? string.Empty : GetSeriesValue(values, i).ToString("0.##", CultureInfo.InvariantCulture)
                });
            }

            return rows.ToArray();
        }

        private static int GetChartLegendColorCount(ExcelChartSnapshot snapshot) {
            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                return snapshot.Data.Categories.Count;
            }

            return snapshot.Data.Series.Count;
        }

        private static PdfCore.PdfTableStyle CreateChartLegendStyle(int colorCount) {
            var style = new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                FontSize = 8.5,
                HeaderFontSize = 8.5,
                CellPaddingX = 4,
                CellPaddingY = 2,
                BorderColor = PdfCore.PdfColor.FromRgb(203, 213, 225),
                HeaderFill = PdfCore.PdfColor.FromRgb(239, 246, 255),
                ColumnWidthWeights = new List<double> { 0.7D, 1.3D },
                AutoFitColumns = false,
                MaxWidth = 300D,
                SpacingAfter = 6
            };

            var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            for (int i = 0; i < colorCount; i++) {
                fills[(i + 1, 0)] = PdfCore.PdfColor.FromOfficeColor(GetChartSeriesColor(i));
            }
            style.CellFills = fills;
            return style;
        }

        private static double GetPositiveMax(IReadOnlyList<ExcelChartSeries> series) {
            double max = 0D;
            foreach (ExcelChartSeries item in series) {
                foreach (double value in item.Values) {
                    if (!double.IsNaN(value) && !double.IsInfinity(value) && value > max) {
                        max = value;
                    }
                }
            }

            return max <= 0D ? 1D : max;
        }

        private static double GetPositiveStackedMax(IReadOnlyList<ExcelChartSeries> series, int categoryCount) {
            double max = 0D;
            for (int i = 0; i < categoryCount; i++) {
                double total = GetPositiveCategoryTotal(series, i);
                if (total > max) {
                    max = total;
                }
            }

            return max <= 0D ? 1D : max;
        }

        private static double GetPositiveCategoryTotal(IReadOnlyList<ExcelChartSeries> series, int categoryIndex) {
            double total = 0D;
            for (int s = 0; s < series.Count; s++) {
                total += Math.Max(0D, GetSeriesValue(series[s], categoryIndex));
            }

            return total;
        }

        private static (double Min, double Max) GetStackedSeriesRange(IReadOnlyList<ExcelChartSeries> series, int categoryCount) {
            double min = 0D;
            double max = 0D;
            for (int category = 0; category < categoryCount; category++) {
                double positive = 0D;
                double negative = 0D;
                for (int s = 0; s < series.Count; s++) {
                    double value = GetSeriesValue(series[s], category);
                    if (value >= 0D) {
                        positive += value;
                    } else {
                        negative += value;
                    }
                }

                if (positive > max) max = positive;
                if (negative < min) min = negative;
            }

            return ExpandFlatRange(min, max);
        }

        private static double GetSeriesValue(ExcelChartSeries series, int index) {
            double value = index >= 0 && index < series.Values.Count ? series.Values[index] : 0D;
            return double.IsNaN(value) || double.IsInfinity(value) ? 0D : value;
        }

        private static double ToPlotY(double value, double max, double plotTop, double plotHeight) {
            double ratio = max <= 0D ? 0D : Math.Max(0D, value) / max;
            if (ratio > 1D) {
                ratio = 1D;
            }

            return plotTop + plotHeight - (plotHeight * ratio);
        }

        private static double ToPlotY(double value, double min, double max, double plotTop, double plotHeight) {
            double range = max - min;
            double ratio = range <= 0D ? 0.5D : (value - min) / range;
            if (ratio < 0D) {
                ratio = 0D;
            } else if (ratio > 1D) {
                ratio = 1D;
            }

            return plotTop + plotHeight - (plotHeight * ratio);
        }

        private static double ToPlotX(double value, double min, double max, double plotLeft, double plotWidth) {
            double range = max - min;
            double ratio = range <= 0D ? 0.5D : (value - min) / range;
            if (ratio < 0D) {
                ratio = 0D;
            } else if (ratio > 1D) {
                ratio = 1D;
            }

            return plotLeft + plotWidth * ratio;
        }

        private static IReadOnlyList<double> GetScatterXValues(IReadOnlyList<string> categories) {
            var values = new double[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (double.TryParse(categories[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    values[i] = value;
                } else {
                    values[i] = i + 1D;
                }
            }

            return values;
        }

        private static IReadOnlyList<OfficePoint> CreateRadarPoints(int count, double centerX, double centerY, double radius) {
            var points = new List<OfficePoint>(count);
            for (int i = 0; i < count; i++) {
                points.Add(CreateRadarPoint(i, count, centerX, centerY, radius));
            }

            return points;
        }

        private static OfficePoint CreateRadarPoint(int index, int count, double centerX, double centerY, double radius) {
            double angle = -Math.PI / 2D + Math.PI * 2D * index / count;
            return new OfficePoint(centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
        }

        private static (double Min, double Max) GetFiniteSeriesRange(IReadOnlyList<ExcelChartSeries> series) {
            bool any = false;
            double min = 0D;
            double max = 0D;
            foreach (ExcelChartSeries item in series) {
                foreach (double value in item.Values) {
                    if (double.IsNaN(value) || double.IsInfinity(value)) {
                        continue;
                    }

                    if (!any) {
                        min = value;
                        max = value;
                        any = true;
                    } else {
                        if (value < min) min = value;
                        if (value > max) max = value;
                    }
                }
            }

            return any ? ExpandFlatRange(min, max) : (0D, 1D);
        }

        private static (double Min, double Max) GetFiniteRange(IReadOnlyList<double> values) {
            bool any = false;
            double min = 0D;
            double max = 0D;
            foreach (double value in values) {
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    continue;
                }

                if (!any) {
                    min = value;
                    max = value;
                    any = true;
                } else {
                    if (value < min) min = value;
                    if (value > max) max = value;
                }
            }

            return any ? ExpandFlatRange(min, max) : (0D, 1D);
        }

        private static (double Min, double Max) ExpandFlatRange(double min, double max) {
            if (max > min) {
                return (min, max);
            }

            double padding = Math.Abs(min) > 1D ? Math.Abs(min) * 0.1D : 1D;
            return (min - padding, max + padding);
        }

        private static OfficeColor GetChartSeriesColor(int index) {
            switch (index % 6) {
                case 0:
                    return OfficeColor.FromRgb(31, 78, 121);
                case 1:
                    return OfficeColor.FromRgb(47, 111, 62);
                case 2:
                    return OfficeColor.FromRgb(184, 90, 35);
                case 3:
                    return OfficeColor.FromRgb(112, 48, 160);
                case 4:
                    return OfficeColor.FromRgb(37, 99, 235);
                default:
                    return OfficeColor.FromRgb(120, 113, 108);
            }
        }

        private static void AddShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, OfficeColor? fill, OfficeColor? stroke, double strokeWidth) {
            shape.FillColor = fill;
            shape.StrokeColor = stroke;
            shape.StrokeWidth = strokeWidth;
            drawing.AddShape(shape, x, y);
        }

        private static void AddPolygonShape(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, double? fillOpacity = null) {
            if (points.Count < 3) {
                return;
            }

            double minX = points[0].X;
            double minY = points[0].Y;
            double maxX = points[0].X;
            double maxY = points[0].Y;
            for (int i = 1; i < points.Count; i++) {
                OfficePoint point = points[i];
                if (point.X < minX) minX = point.X;
                if (point.Y < minY) minY = point.Y;
                if (point.X > maxX) maxX = point.X;
                if (point.Y > maxY) maxY = point.Y;
            }

            if (maxX <= minX || maxY <= minY) {
                return;
            }

            OfficeShape shape = OfficeShape.Polygon(points);
            shape.FillOpacity = fillOpacity;
            AddShape(drawing, shape, minX, minY, fill, stroke, strokeWidth);
        }

        private static void AddPointLine(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth) {
            for (int i = 1; i < points.Count; i++) {
                OfficePoint previous = points[i - 1];
                OfficePoint current = points[i];
                if (previous.Equals(current)) {
                    continue;
                }

                double minX = Math.Min(previous.X, current.X);
                double minY = Math.Min(previous.Y, current.Y);
                AddShape(
                    drawing,
                    OfficeShape.Line(previous.X - minX, previous.Y - minY, current.X - minX, current.Y - minY),
                    minX,
                    minY,
                    null,
                    color,
                    strokeWidth);
            }
        }

        private static bool IsColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnClustered
                   || type == ExcelChartType.ColumnStacked
                   || type == ExcelChartType.ColumnStacked100
                   || type == ExcelChartType.Column3DClustered
                   || type == ExcelChartType.Column3DStacked
                   || type == ExcelChartType.Column3DStacked100;
        }

        private static bool IsBarChart(ExcelChartType type) {
            return type == ExcelChartType.BarClustered
                   || type == ExcelChartType.BarStacked
                   || type == ExcelChartType.BarStacked100
                   || type == ExcelChartType.Bar3DClustered
                   || type == ExcelChartType.Bar3DStacked
                   || type == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsLineChart(ExcelChartType type) {
            return type == ExcelChartType.Line
                   || type == ExcelChartType.LineStacked
                   || type == ExcelChartType.LineStacked100
                   || type == ExcelChartType.Line3D;
        }

        private static bool IsAreaChart(ExcelChartType type) {
            return type == ExcelChartType.Area
                   || type == ExcelChartType.AreaStacked
                   || type == ExcelChartType.AreaStacked100
                   || type == ExcelChartType.Area3D
                   || type == ExcelChartType.Area3DStacked
                   || type == ExcelChartType.Area3DStacked100;
        }

        private static bool IsScatterChart(ExcelChartType type) {
            return type == ExcelChartType.Scatter;
        }

        private static bool IsRadarChart(ExcelChartType type) {
            return type == ExcelChartType.Radar;
        }

        private static bool IsStackedAreaChart(ExcelChartType type) {
            return type == ExcelChartType.AreaStacked
                   || type == ExcelChartType.Area3DStacked;
        }

        private static bool IsPercentStackedAreaChart(ExcelChartType type) {
            return type == ExcelChartType.AreaStacked100
                   || type == ExcelChartType.Area3DStacked100;
        }

        private static bool IsStackedBarOrColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnStacked
                   || type == ExcelChartType.Column3DStacked
                   || type == ExcelChartType.BarStacked
                   || type == ExcelChartType.Bar3DStacked;
        }

        private static bool IsPercentStackedBarOrColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnStacked100
                   || type == ExcelChartType.Column3DStacked100
                   || type == ExcelChartType.BarStacked100
                   || type == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsPieChart(ExcelChartType type) {
            return type == ExcelChartType.Pie
                   || type == ExcelChartType.Pie3D
                   || type == ExcelChartType.PieOfPie
                   || type == ExcelChartType.BarOfPie;
        }

        private static bool IsDoughnutChart(ExcelChartType type) {
            return type == ExcelChartType.Doughnut;
        }

        private static ExcelSheet? GetWorkbookSheet(ExcelDocument document, string sheetName) {
            foreach (ExcelSheet sheet in document.Sheets) {
                if (string.Equals(sheet.Name, sheetName, StringComparison.Ordinal)) {
                    return sheet;
                }
            }

            return null;
        }

        private static string GetExportRange(ExcelSheetReader sheet, ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) {
            string? printArea = GetWorksheetPrintArea(workbookSheet, options);
            if (!string.IsNullOrWhiteSpace(printArea)) {
                if (ContainsMultiplePrintAreas(printArea!)) {
                    AddWarning(
                        options,
                        sheet.Name,
                        "WorksheetPrintArea",
                        "Multi-area worksheet print areas are not supported by the first-party PDF exporter; exporting the worksheet used range instead.");
                    return sheet.GetUsedRangeA1();
                }

                return NormalizeA1Range(printArea!);
            }

            return sheet.GetUsedRangeA1();
        }

        private static bool ContainsMultiplePrintAreas(string printArea) {
            bool inQuotedSheetName = false;
            for (int i = 0; i < printArea.Length; i++) {
                char current = printArea[i];
                if (current == '\'') {
                    if (inQuotedSheetName && i + 1 < printArea.Length && printArea[i + 1] == '\'') {
                        i++;
                    } else {
                        inQuotedSheetName = !inQuotedSheetName;
                    }
                } else if (current == ',' && !inQuotedSheetName) {
                    return true;
                }
            }

            return false;
        }

        private static bool HasWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) =>
            !string.IsNullOrWhiteSpace(GetWorksheetPrintArea(workbookSheet, options));

        private static string? GetWorksheetPrintArea(ExcelSheet? workbookSheet, ExcelPdfSaveOptions options) =>
            options.UseWorksheetPrintAreas && workbookSheet != null ? workbookSheet.GetPrintArea() : null;

        private static SheetExportData ReadSheetExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string exportRange, ExcelPdfSaveOptions options) {
            string normalizedRange = NormalizeA1Range(exportRange);
            A1.TryParseRange(normalizedRange, out int rangeFirstRow, out int rangeFirstColumn, out _, out int rangeLastColumn);
            RangeExportData bodyRange = ReadRangeExportData(sheet, workbookSheet, normalizedRange, options);
            object?[,] values = bodyRange.Values;
            ExcelCellStyleSnapshot?[,]? styles = bodyRange.Styles;
            ExcelHyperlinkSnapshot?[,]? hyperlinks = bodyRange.Hyperlinks;
            string?[,]? cellReferences = bodyRange.CellReferences;
            MergeLayoutData? mergedCells = bodyRange.MergedCells;
            ColumnLayoutData? columnWidths = bodyRange.ColumnWidths;
            RowLayoutData? rowHeights = bodyRange.RowHeights;
            int headerRows = options.HeaderRowCount;
            if (!options.UseWorksheetPrintTitleRows || workbookSheet == null) {
                return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
            }

            ExcelPrintTitles titles = workbookSheet.GetPrintTitles();
            if (!titles.HasRows) {
                return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
            }

            int firstTitleRow = titles.FirstRow!.Value;
            int lastTitleRow = titles.LastRow!.Value;
            if (firstTitleRow < rangeFirstRow) {
                int prependedLastTitleRow = Math.Min(lastTitleRow, rangeFirstRow - 1);
                string titleRange = ToA1Range(firstTitleRow, rangeFirstColumn, prependedLastTitleRow, rangeLastColumn);
                RangeExportData titleRangeData = ReadRangeExportData(sheet, workbookSheet, titleRange, options);
                int prependedRowCount = titleRangeData.Values.GetLength(0);
                int bodyRowCount = values.GetLength(0);
                int columnCount = values.GetLength(1);
                object?[,] prependedValues = PrependRows(titleRangeData.Values, values);
                ExcelCellStyleSnapshot?[,]? prependedStyles = PrependRows(titleRangeData.Styles, styles, prependedRowCount, bodyRowCount, columnCount);
                ExcelHyperlinkSnapshot?[,]? prependedHyperlinks = PrependRows(titleRangeData.Hyperlinks, hyperlinks, prependedRowCount, bodyRowCount, columnCount);
                string?[,]? prependedCellReferences = PrependRows(titleRangeData.CellReferences, cellReferences, prependedRowCount, bodyRowCount, columnCount);
                MergeLayoutData? prependedMergedCells = PrependRows(titleRangeData.MergedCells, mergedCells, prependedRowCount, bodyRowCount, columnCount);
                RowLayoutData? prependedRowHeights = PrependRows(titleRangeData.RowHeights, rowHeights, prependedRowCount, bodyRowCount);
                int overlappingTitleRows = lastTitleRow >= rangeFirstRow
                    ? Math.Min(bodyRowCount, lastTitleRow - rangeFirstRow + 1)
                    : 0;
                return CreateSheetExportData(
                    workbookSheet,
                    prependedValues,
                    prependedStyles,
                    prependedHyperlinks,
                    prependedCellReferences,
                    prependedMergedCells,
                    columnWidths,
                    prependedRowHeights,
                    Math.Max(headerRows, prependedRowCount + overlappingTitleRows),
                    rangeFirstRow,
                    options);
            }

            if (firstTitleRow <= rangeFirstRow && lastTitleRow >= rangeFirstRow) {
                int titleRowsInsideRange = Math.Min(values.GetLength(0), lastTitleRow - rangeFirstRow + 1);
                headerRows = Math.Max(headerRows, titleRowsInsideRange);
            }

            return CreateSheetExportData(workbookSheet, values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, rangeFirstRow, options);
        }

        private static SheetExportData CreateSheetExportData(ExcelSheet? workbookSheet, object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRows, int firstBodyRowNumber, ExcelPdfSaveOptions options) {
            ConditionalFillData? conditionalFills = ReadConditionalFillData(
                workbookSheet,
                values,
                cellReferences,
                options.UseWorksheetCellStyles);

            return new SheetExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights, headerRows, firstBodyRowNumber, conditionalFills);
        }

        private static RangeExportData ReadRangeExportData(ExcelSheetReader sheet, ExcelSheet? workbookSheet, string normalizedRange, ExcelPdfSaveOptions options) {
            object?[,] rawValues = sheet.ReadRange(normalizedRange);
            VisibilityLayoutData? visibility = ReadVisibilityLayoutData(
                workbookSheet,
                normalizedRange,
                rawValues.GetLength(0),
                rawValues.GetLength(1),
                options.RespectWorksheetHiddenRowsAndColumns);
            object?[,] values = FilterValues(rawValues, visibility);
            int rowCount = values.GetLength(0);
            int columnCount = values.GetLength(1);
            string?[,]? cellReferences = ReadCellReferenceData(
                normalizedRange,
                rowCount,
                columnCount,
                visibility);
            ExcelCellStyleSnapshot?[,]? styles = ReadCellStyleData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetCellStyles,
                visibility);
            ExcelHyperlinkSnapshot?[,]? hyperlinks = ReadHyperlinkData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetHyperlinks,
                visibility);
            MergeLayoutData? mergedCells = ReadMergeLayoutData(
                workbookSheet,
                normalizedRange,
                rowCount,
                columnCount,
                options.UseWorksheetMergedCells,
                visibility);
            ColumnLayoutData? columnWidths = ReadColumnLayoutData(
                workbookSheet,
                normalizedRange,
                columnCount,
                options.UseWorksheetColumnWidths,
                visibility);
            RowLayoutData? rowHeights = ReadRowLayoutData(
                workbookSheet,
                normalizedRange,
                rowCount,
                options.UseWorksheetRowHeights,
                visibility);

            return new RangeExportData(values, styles, hyperlinks, cellReferences, mergedCells, columnWidths, rowHeights);
        }

        private static ConditionalFillData? ReadConditionalFillData(ExcelSheet? workbookSheet, object?[,] values, string?[,]? cellReferences, bool enabled) {
            if (!enabled || workbookSheet == null || cellReferences == null) {
                return null;
            }

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = workbookSheet.GetConditionalFormattingRules();
            if (rules.Count == 0) {
                return null;
            }

            var fills = new Dictionary<(int Row, int Column), string>();
            var dataBars = new Dictionary<(int Row, int Column), ConditionalDataBarCell>();
            var icons = new Dictionary<(int Row, int Column), ConditionalIconCell>();
            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "ColorScale", StringComparison.OrdinalIgnoreCase) && rule.ColorScaleColors.Count >= 2)
                .OrderByDescending(rule => rule.Priority)) {
                if (!TryGetRgb(rule.ColorScaleColors[0], out byte startR, out byte startG, out byte startB) ||
                    !TryGetRgb(rule.ColorScaleColors[rule.ColorScaleColors.Count - 1], out byte endR, out byte endG, out byte endB)) {
                    continue;
                }

                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    double ratio = max <= min ? 0.5D : Math.Max(0D, Math.Min(1D, (candidate.Value - min) / (max - min)));
                    fills[(candidate.Row, candidate.Column)] = InterpolateRgbHex(startR, startG, startB, endR, endG, endB, ratio);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "DataBar", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.DataBarColor))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    double ratio = max <= min ? 1D : Math.Max(0D, Math.Min(1D, (candidate.Value - min) / (max - min)));
                    dataBars[(candidate.Row, candidate.Column)] = new ConditionalDataBarCell(rule.DataBarColor!, ratio);
                }
            }

            foreach (ExcelConditionalFormattingInfo rule in rules
                .Where(rule => string.Equals(rule.Type, "IconSet", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(rule.IconSet))
                .OrderByDescending(rule => rule.Priority)) {
                var candidates = new List<(int Row, int Column, double Value)>();
                for (int row = 0; row < values.GetLength(0); row++) {
                    for (int column = 0; column < values.GetLength(1); column++) {
                        string? cellReference = cellReferences[row, column];
                        if (!string.IsNullOrWhiteSpace(cellReference) &&
                            IsCellReferenceInReferenceList(cellReference!, rule.Range) &&
                            TryGetConditionalNumericValue(values[row, column], out double numericValue)) {
                            candidates.Add((row, column, numericValue));
                        }
                    }
                }

                if (candidates.Count == 0) {
                    continue;
                }

                int iconCount = GetExcelIconSetCount(rule.IconSet!);
                double min = candidates.Min(candidate => candidate.Value);
                double max = candidates.Max(candidate => candidate.Value);
                foreach (var candidate in candidates) {
                    int bucket = GetExcelIconSetBucket(candidate.Value, min, max, iconCount);
                    if (rule.IconSetReverse) {
                        bucket = iconCount - 1 - bucket;
                    }

                    icons[(candidate.Row, candidate.Column)] = MapExcelIconSetCell(rule.IconSet!, bucket, iconCount);
                }
            }

            return fills.Count == 0 && dataBars.Count == 0 && icons.Count == 0 ? null : new ConditionalFillData(fills, dataBars, icons);
        }

        private static int GetExcelIconSetCount(string iconSet) {
            if (iconSet.StartsWith("Three", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("3", StringComparison.Ordinal)) {
                return 3;
            }

            if (iconSet.StartsWith("Four", StringComparison.OrdinalIgnoreCase) ||
                iconSet.StartsWith("4", StringComparison.Ordinal)) {
                return 4;
            }

            return 5;
        }

        private static int GetExcelIconSetBucket(double value, double min, double max, int iconCount) {
            if (iconCount <= 1 || max <= min) {
                return iconCount - 1;
            }

            double ratio = Math.Max(0D, Math.Min(1D, (value - min) / (max - min)));
            return Math.Max(0, Math.Min(iconCount - 1, (int)Math.Floor(ratio * iconCount)));
        }

        private static ConditionalIconCell MapExcelIconSetCell(string iconSet, int bucket, int iconCount) {
            string normalized = iconSet.ToLowerInvariant();
            bool trafficLights = normalized.IndexOf("traffic", StringComparison.Ordinal) >= 0;
            bool arrows = normalized.IndexOf("arrow", StringComparison.Ordinal) >= 0;
            bool symbols = normalized.IndexOf("symbol", StringComparison.Ordinal) >= 0 || normalized.IndexOf("sign", StringComparison.Ordinal) >= 0 || normalized.IndexOf("indicator", StringComparison.Ordinal) >= 0;

            if (trafficLights) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
            }

            if (arrows) {
                if (bucket == 0) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleDown, PdfCore.PdfColor.FromRgb(192, 80, 77));
                }

                if (bucket >= iconCount - 1) {
                    return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(99, 155, 71));
                }

                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleRight, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            if (symbols && bucket == 0) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Diamond, PdfCore.PdfColor.FromRgb(192, 80, 77));
            }

            if (symbols && bucket >= iconCount - 1) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, PdfCore.PdfColor.FromRgb(99, 155, 71));
            }

            if (symbols) {
                return new ConditionalIconCell(PdfCore.PdfCellIconKind.TriangleUp, PdfCore.PdfColor.FromRgb(255, 192, 0));
            }

            return new ConditionalIconCell(PdfCore.PdfCellIconKind.Circle, GetExcelIconBucketColor(bucket, iconCount));
        }

        private static PdfCore.PdfColor GetExcelIconBucketColor(int bucket, int iconCount) {
            if (bucket <= 0) {
                return PdfCore.PdfColor.FromRgb(192, 80, 77);
            }

            if (bucket >= iconCount - 1) {
                return PdfCore.PdfColor.FromRgb(99, 155, 71);
            }

            return PdfCore.PdfColor.FromRgb(255, 192, 0);
        }

        private static bool IsCellReferenceInReferenceList(string cellReference, string referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            (int Row, int Col) cell = A1.ParseCellRef(NormalizeCellReference(cellReference));
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            foreach (string rawToken in referenceList.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                string token = StripSheetPrefix(rawToken).Replace("$", string.Empty);
                if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    if (cell.Row >= firstRow && cell.Row <= lastRow && cell.Col >= firstColumn && cell.Col <= lastColumn) {
                        return true;
                    }
                } else {
                    (int Row, int Col) singleCell = A1.ParseCellRef(token);
                    if (singleCell.Row == cell.Row && singleCell.Col == cell.Col) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryGetConditionalNumericValue(object? value, out double numericValue) {
            if (value is DateTime dateTime) {
                numericValue = dateTime.ToOADate();
                return true;
            }

            if (value is IConvertible convertible) {
                try {
                    numericValue = convertible.ToDouble(CultureInfo.InvariantCulture);
                    return !double.IsNaN(numericValue) && !double.IsInfinity(numericValue);
                } catch (FormatException) {
                } catch (InvalidCastException) {
                } catch (OverflowException) {
                }
            }

            numericValue = 0D;
            return false;
        }

        private static bool TryGetRgb(string value, out byte r, out byte g, out byte b) {
            string normalized = value.Trim().TrimStart('#');
            if (normalized.Length == 8) {
                normalized = normalized.Substring(2);
            }

            if (normalized.Length != 6 ||
                !byte.TryParse(normalized.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r) ||
                !byte.TryParse(normalized.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g) ||
                !byte.TryParse(normalized.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b)) {
                r = 0;
                g = 0;
                b = 0;
                return false;
            }

            return true;
        }

        private static string InterpolateRgbHex(byte startR, byte startG, byte startB, byte endR, byte endG, byte endB, double ratio) {
            byte r = InterpolateByte(startR, endR, ratio);
            byte g = InterpolateByte(startG, endG, ratio);
            byte b = InterpolateByte(startB, endB, ratio);
            return r.ToString("X2", CultureInfo.InvariantCulture) +
                g.ToString("X2", CultureInfo.InvariantCulture) +
                b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static byte InterpolateByte(byte start, byte end, double ratio) {
            return (byte)Math.Max(0, Math.Min(255, (int)Math.Round(start + ((end - start) * ratio), MidpointRounding.AwayFromZero)));
        }

        private static VisibilityLayoutData? ReadVisibilityLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelRowSnapshot> rowDefinitions = workbookSheet.GetRowDefinitions();
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions = workbookSheet.GetColumnDefinitions();
            if (!rowDefinitions.Any(row => row.Hidden) && !columnDefinitions.Any(column => column.Hidden)) {
                return null;
            }

            var rowOffsets = new List<int>(rowCount);
            for (int row = 0; row < rowCount; row++) {
                if (!IsWorksheetRowHidden(rowDefinitions, firstRow + row)) {
                    rowOffsets.Add(row);
                }
            }

            var columnOffsets = new List<int>(columnCount);
            for (int column = 0; column < columnCount; column++) {
                if (!IsWorksheetColumnHidden(columnDefinitions, firstColumn + column)) {
                    columnOffsets.Add(column);
                }
            }

            if (rowOffsets.Count == rowCount && columnOffsets.Count == columnCount) {
                return null;
            }

            return new VisibilityLayoutData(rowOffsets, columnOffsets, rowCount, columnCount);
        }

        private static string?[,]? ReadCellReferenceData(string normalizedRange, int rowCount, int columnCount, VisibilityLayoutData? visibility = null) {
            if (rowCount == 0 || columnCount == 0 ||
                !A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            var references = new string?[rowCount, columnCount];
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    references[row, column] = A1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
                }
            }

            return references;
        }

        private static object?[,] FilterValues(object?[,] values, VisibilityLayoutData? visibility) {
            if (visibility == null) {
                return values;
            }

            var result = new object?[visibility.RowOffsets.Count, visibility.ColumnOffsets.Count];
            for (int row = 0; row < visibility.RowOffsets.Count; row++) {
                for (int column = 0; column < visibility.ColumnOffsets.Count; column++) {
                    result[row, column] = values[visibility.RowOffsets[row], visibility.ColumnOffsets[column]];
                }
            }

            return result;
        }

        private static bool IsWorksheetRowHidden(IReadOnlyList<ExcelRowSnapshot> rowDefinitions, int rowIndex) {
            for (int i = rowDefinitions.Count - 1; i >= 0; i--) {
                ExcelRowSnapshot definition = rowDefinitions[i];
                if (definition.Index == rowIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static bool IsWorksheetColumnHidden(IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, int columnIndex) {
            for (int i = columnDefinitions.Count - 1; i >= 0; i--) {
                ExcelColumnSnapshot definition = columnDefinitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }

        private static ExcelCellStyleSnapshot?[,]? ReadCellStyleData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            ExcelCellStyleSnapshot?[,] styles = new ExcelCellStyleSnapshot?[rowCount, columnCount];
            bool hasAnyStyle = false;
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    int sourceRow = visibility?.RowOffsets[row] ?? row;
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    ExcelCellStyleSnapshot style = workbookSheet.GetCellStyle(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (style.HasPdfVisualStyle) {
                        styles[row, column] = style;
                        hasAnyStyle = true;
                    }
                }
            }

            return hasAnyStyle ? styles : null;
        }

        private static ExcelHyperlinkSnapshot?[,]? ReadHyperlinkData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> worksheetHyperlinks = workbookSheet.GetHyperlinks();
            if (worksheetHyperlinks.Count == 0) {
                return null;
            }

            var links = new ExcelHyperlinkSnapshot?[rowCount, columnCount];
            bool hasAnyLink = false;
            for (int row = 0; row < rowCount; row++) {
                int sourceRow = visibility?.RowOffsets[row] ?? row;
                for (int column = 0; column < columnCount; column++) {
                    int sourceColumn = visibility?.ColumnOffsets[column] ?? column;
                    string reference = A1.CellReference(firstRow + sourceRow, firstColumn + sourceColumn);
                    if (TryGetHyperlink(worksheetHyperlinks, reference, out ExcelHyperlinkSnapshot? hyperlink) &&
                        IsSupportedPdfHyperlink(hyperlink, workbookSheet.Name)) {
                        links[row, column] = hyperlink;
                        hasAnyLink = true;
                    }
                }
            }

            return hasAnyLink ? links : null;
        }

        private static bool IsSupportedPdfHyperlink(ExcelHyperlinkSnapshot hyperlink, string currentSheetName) {
            if (hyperlink.IsExternal) {
                return Uri.TryCreate(hyperlink.Target, UriKind.Absolute, out _);
            }

            return TryParseInternalSheetName(hyperlink.Target, currentSheetName, out _);
        }

        private static bool TryGetHyperlink(IReadOnlyDictionary<string, ExcelHyperlinkSnapshot> hyperlinks, string cellReference, out ExcelHyperlinkSnapshot hyperlink) {
            if (hyperlinks.TryGetValue(cellReference, out ExcelHyperlinkSnapshot? direct)) {
                hyperlink = direct;
                return true;
            }

            foreach (KeyValuePair<string, ExcelHyperlinkSnapshot> entry in hyperlinks) {
                (int Row, int Col) cell = A1.ParseCellRef(cellReference);
                if (A1.TryParseRange(entry.Key, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) &&
                    cell.Row >= firstRow &&
                    cell.Row <= lastRow &&
                    cell.Col >= firstColumn &&
                    cell.Col <= lastColumn) {
                    hyperlink = entry.Value;
                    return true;
                }
            }

            hyperlink = null!;
            return false;
        }

        private static ColumnLayoutData? ReadColumnLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out _, out int firstColumn, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions = workbookSheet.GetColumnDefinitions();
            if (columnDefinitions.Count == 0) {
                return null;
            }

            var weights = new List<double>(columnCount);
            bool hasCustomWidth = false;
            double totalWidth = 0D;
            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++) {
                int sourceColumnOffset = visibility?.ColumnOffsets[columnOffset] ?? columnOffset;
                int absoluteColumn = firstColumn + sourceColumnOffset;
                double width = GetWorksheetColumnWidth(columnDefinitions, absoluteColumn, out bool customWidth);
                weights.Add(width);
                totalWidth += width;
                hasCustomWidth |= customWidth;
            }

            return hasCustomWidth ? new ColumnLayoutData(weights, totalWidth * 5.25D) : null;
        }

        private static double GetWorksheetColumnWidth(IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, int columnIndex, out bool customWidth) {
            for (int i = columnDefinitions.Count - 1; i >= 0; i--) {
                ExcelColumnSnapshot definition = columnDefinitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    customWidth = definition.CustomWidth && definition.Width.HasValue && definition.Width.Value > 0D;
                    return customWidth ? definition.Width!.Value : 8.43D;
                }
            }

            customWidth = false;
            return 8.43D;
        }

        private static RowLayoutData? ReadRowLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out _, out _, out _)) {
                return null;
            }

            IReadOnlyList<ExcelRowSnapshot> rowDefinitions = workbookSheet.GetRowDefinitions();
            if (rowDefinitions.Count == 0) {
                return null;
            }

            var minHeights = new List<double?>(rowCount);
            bool hasCustomHeight = false;
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++) {
                int sourceRowOffset = visibility?.RowOffsets[rowOffset] ?? rowOffset;
                int absoluteRow = firstRow + sourceRowOffset;
                double? height = GetWorksheetRowHeight(rowDefinitions, absoluteRow);
                minHeights.Add(height);
                hasCustomHeight |= height.HasValue;
            }

            return hasCustomHeight ? new RowLayoutData(minHeights) : null;
        }

        private static double? GetWorksheetRowHeight(IReadOnlyList<ExcelRowSnapshot> rowDefinitions, int rowIndex) {
            for (int i = rowDefinitions.Count - 1; i >= 0; i--) {
                ExcelRowSnapshot definition = rowDefinitions[i];
                if (definition.Index == rowIndex) {
                    return definition.CustomHeight && definition.Height.HasValue && definition.Height.Value > 0D
                        ? definition.Height.Value
                        : null;
                }
            }

            return null;
        }

        private static MergeLayoutData? ReadMergeLayoutData(ExcelSheet? workbookSheet, string normalizedRange, int rowCount, int columnCount, bool enabled, VisibilityLayoutData? visibility = null) {
            if (!enabled || workbookSheet == null || rowCount == 0 || columnCount == 0) {
                return null;
            }

            if (!A1.TryParseRange(normalizedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return null;
            }

            var layout = new MergeLayoutData(rowCount, columnCount);
            foreach (ExcelMergedRangeSnapshot mergedRange in workbookSheet.GetMergedRanges()) {
                if (mergedRange.StartRow < firstRow ||
                    mergedRange.StartColumn < firstColumn ||
                    mergedRange.EndRow > lastRow ||
                    mergedRange.EndColumn > lastColumn) {
                    continue;
                }

                List<int> visibleRows = MapVisibleOffsets(mergedRange.StartRow - firstRow, mergedRange.EndRow - firstRow, visibility?.RowOffsets);
                List<int> visibleColumns = MapVisibleOffsets(mergedRange.StartColumn - firstColumn, mergedRange.EndColumn - firstColumn, visibility?.ColumnOffsets);
                if (visibleRows.Count == 0 || visibleColumns.Count == 0) {
                    continue;
                }

                int relativeRow = visibleRows[0];
                int relativeColumn = visibleColumns[0];
                int rowSpan = visibleRows.Count;
                int columnSpan = visibleColumns.Count;
                if (rowSpan > 1 || columnSpan > 1) {
                    layout.SetSpan(relativeRow, relativeColumn, rowSpan, columnSpan);
                }
            }

            return layout.HasAny ? layout : null;
        }

        private static List<int> MapVisibleOffsets(int firstSourceOffset, int lastSourceOffset, IReadOnlyList<int>? visibleOffsets) {
            if (visibleOffsets == null) {
                var all = new List<int>(lastSourceOffset - firstSourceOffset + 1);
                for (int offset = firstSourceOffset; offset <= lastSourceOffset; offset++) {
                    all.Add(offset);
                }

                return all;
            }

            var mapped = new List<int>();
            for (int index = 0; index < visibleOffsets.Count; index++) {
                int sourceOffset = visibleOffsets[index];
                if (sourceOffset >= firstSourceOffset && sourceOffset <= lastSourceOffset) {
                    mapped.Add(index);
                }
            }

            return mapped;
        }

        private static object?[,] PrependRows(object?[,] topRows, object?[,] bodyRows) {
            int topRowCount = topRows.GetLength(0);
            int bodyRowCount = bodyRows.GetLength(0);
            int columnCount = bodyRows.GetLength(1);
            var result = new object?[topRowCount + bodyRowCount, columnCount];
            CopyRows(topRows, result, 0, columnCount);
            CopyRows(bodyRows, result, topRowCount, columnCount);
            return result;
        }

        private static ExcelCellStyleSnapshot?[,]? PrependRows(ExcelCellStyleSnapshot?[,]? topRows, ExcelCellStyleSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelCellStyleSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static ExcelHyperlinkSnapshot?[,]? PrependRows(ExcelHyperlinkSnapshot?[,]? topRows, ExcelHyperlinkSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelHyperlinkSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static string?[,]? PrependRows(string?[,]? topRows, string?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new string?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static MergeLayoutData? PrependRows(MergeLayoutData? topRows, MergeLayoutData? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new MergeLayoutData(topRowCount + bodyRowCount, columnCount);
            topRows?.CopyTo(result, 0);
            bodyRows?.CopyTo(result, topRowCount);
            return result.HasAny ? result : null;
        }

        private static RowLayoutData? PrependRows(RowLayoutData? topRows, RowLayoutData? bodyRows, int topRowCount, int bodyRowCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var minHeights = new List<double?>(topRowCount + bodyRowCount);
            if (topRows != null) {
                minHeights.AddRange(topRows.MinHeights);
            } else {
                for (int row = 0; row < topRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            if (bodyRows != null) {
                minHeights.AddRange(bodyRows.MinHeights);
            } else {
                for (int row = 0; row < bodyRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            return minHeights.Any(height => height.HasValue) ? new RowLayoutData(minHeights) : null;
        }

        private static void CopyRows(object?[,] source, object?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelCellStyleSnapshot?[,] source, ExcelCellStyleSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelHyperlinkSnapshot?[,] source, ExcelHyperlinkSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(string?[,] source, string?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static string NormalizeA1Range(string range) {
            string withoutSheet = StripSheetPrefix(range).Replace("$", string.Empty);
            if (!A1.TryParseRange(withoutSheet, out int r1, out int c1, out int r2, out int c2)) {
                (int Row, int Col) cell = A1.ParseCellRef(withoutSheet);
                if (cell.Row <= 0 || cell.Col <= 0) {
                    throw new ArgumentException("Excel PDF export range must be a valid A1 range.", nameof(range));
                }

                r1 = r2 = cell.Row;
                c1 = c2 = cell.Col;
            }

            return ToA1Range(r1, c1, r2, c2);
        }

        private static string StripSheetPrefix(string range) {
            int separator = range.LastIndexOf('!');
            return separator >= 0 ? range.Substring(separator + 1) : range;
        }

        private static string ToA1Range(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            string start = A1.ColumnIndexToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
            string end = A1.ColumnIndexToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
            return start + ":" + end;
        }

        private static IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>> CreateWorksheetImageMap(WorksheetPdfExportPlan plan) {
            if (!plan.HasTable || plan.Images.Count == 0 || plan.ExportData.CellReferences == null) {
                return new Dictionary<string, IReadOnlyList<WorksheetImageExportData>>(StringComparer.Ordinal);
            }

            var attachableCellReferences = new HashSet<string>(StringComparer.Ordinal);
            int rows = Math.Min(plan.ExportedRows, plan.ExportData.CellReferences.GetLength(0));
            int columns = plan.ExportData.CellReferences.GetLength(1);
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    if (plan.ExportData.MergedCells?.IsContinuation(row, column) == true) {
                        continue;
                    }

                    string? cellReference = plan.ExportData.CellReferences[row, column];
                    if (!string.IsNullOrWhiteSpace(cellReference)) {
                        attachableCellReferences.Add(NormalizeCellReference(cellReference!));
                    }
                }
            }

            var imagesByCellReference = new Dictionary<string, List<WorksheetImageExportData>>(StringComparer.Ordinal);
            foreach (WorksheetImageExportData image in plan.Images) {
                string cellReference = NormalizeCellReference(image.CellReference);
                if (!attachableCellReferences.Contains(cellReference)) {
                    continue;
                }

                if (!imagesByCellReference.TryGetValue(cellReference, out List<WorksheetImageExportData>? images)) {
                    images = new List<WorksheetImageExportData>();
                    imagesByCellReference[cellReference] = images;
                }

                images.Add(image);
            }

            var result = new Dictionary<string, IReadOnlyList<WorksheetImageExportData>>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, List<WorksheetImageExportData>> item in imagesByCellReference) {
                result[item.Key] = item.Value;
            }

            return result;
        }

        private static ISet<string>? CreateExportedCellReferenceSet(string?[,]? cellReferences, int exportedRows) {
            if (cellReferences == null) {
                return null;
            }

            var exported = new HashSet<string>(StringComparer.Ordinal);
            int rows = Math.Min(exportedRows, cellReferences.GetLength(0));
            int columns = cellReferences.GetLength(1);
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < columns; column++) {
                    string? reference = cellReferences[row, column];
                    if (!string.IsNullOrWhiteSpace(reference)) {
                        exported.Add(NormalizeCellReference(reference!));
                    }
                }
            }

            return exported;
        }

        private static IReadOnlyList<WorksheetImageExportData> FilterImagesByExportedCells(IReadOnlyList<WorksheetImageExportData> images, ISet<string>? exportedCellReferences, bool enabled) {
            if (!enabled || images.Count == 0 || exportedCellReferences == null) {
                return images;
            }

            return images
                .Where(image => exportedCellReferences.Contains(NormalizeCellReference(image.CellReference)))
                .ToList();
        }

        private static IReadOnlyList<WorksheetChartExportData> FilterChartsByExportedCells(IReadOnlyList<WorksheetChartExportData> charts, ISet<string>? exportedCellReferences, bool enabled) {
            if (!enabled || charts.Count == 0 || exportedCellReferences == null) {
                return charts;
            }

            return charts
                .Where(chart => exportedCellReferences.Contains(A1.CellReference(chart.Snapshot.RowIndex, chart.Snapshot.ColumnIndex)))
                .ToList();
        }

        private static string NormalizeCellReference(string cellReference) {
            return cellReference.Replace("$", string.Empty).ToUpperInvariant();
        }

        private static IReadOnlyList<TableChunk> CreateTableChunks(WorksheetPdfExportPlan plan, ExcelPdfSaveOptions options, int exportedColumns) {
            IReadOnlyList<TableAxisChunk> rowChunks = CreateTableAxisChunks(
                plan.ExportedRows,
                options.UseWorksheetPageBreaks ? GetManualRowBreakOffsets(plan) : new List<int>());
            IReadOnlyList<TableAxisChunk> columnChunks = CreateTableAxisChunks(
                exportedColumns,
                options.UseWorksheetPageBreaks ? GetManualColumnBreakOffsets(plan) : new List<int>());
            int headerRowCount = Math.Min(plan.ExportData.HeaderRowCount, plan.ExportedRows);

            var chunks = new List<TableChunk>(rowChunks.Count * columnChunks.Count);
            foreach (TableAxisChunk rowChunk in rowChunks) {
                IReadOnlyList<int> rowIndexes = CreateChunkRowIndexes(rowChunk, headerRowCount);
                int chunkHeaderRows = Math.Min(headerRowCount, rowIndexes.Count);
                foreach (TableAxisChunk columnChunk in columnChunks) {
                    chunks.Add(new TableChunk(rowIndexes, chunkHeaderRows, columnChunk.Start, columnChunk.Count));
                }
            }

            return chunks;
        }

        private static IReadOnlyList<int> CreateChunkRowIndexes(TableAxisChunk rowChunk, int headerRowCount) {
            var indexes = new List<int>(rowChunk.Count + headerRowCount);
            if (rowChunk.Start > 0 && headerRowCount > 0) {
                for (int row = 0; row < headerRowCount; row++) {
                    indexes.Add(row);
                }
            }

            int end = rowChunk.Start + rowChunk.Count;
            for (int row = rowChunk.Start; row < end; row++) {
                if (row < headerRowCount && indexes.Contains(row)) {
                    continue;
                }

                indexes.Add(row);
            }

            return indexes;
        }

        private static IReadOnlyList<TableAxisChunk> CreateTableAxisChunks(int itemCount, IReadOnlyList<int> breakOffsets) {
            if (itemCount <= 0 || breakOffsets.Count == 0) {
                return new[] { new TableAxisChunk(0, itemCount) };
            }

            var chunks = new List<TableAxisChunk>();
            int start = 0;
            foreach (int breakOffset in breakOffsets) {
                if (breakOffset <= start || breakOffset >= itemCount) {
                    continue;
                }

                chunks.Add(new TableAxisChunk(start, breakOffset - start));
                start = breakOffset;
            }

            if (start < itemCount) {
                chunks.Add(new TableAxisChunk(start, itemCount - start));
            }

            return chunks.Count == 0 ? new[] { new TableAxisChunk(0, itemCount) } : chunks;
        }

        private static List<int> GetManualRowBreakOffsets(WorksheetPdfExportPlan plan) {
            var offsets = new SortedSet<int>();
            string?[,]? references = plan.ExportData.CellReferences;
            if (references == null) {
                return offsets.ToList();
            }

            int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
            foreach (int breakRow in plan.ManualRowBreaks) {
                if (breakRow < plan.ExportData.FirstBodyRowNumber) {
                    continue;
                }

                for (int row = 0; row < rows; row++) {
                    int originalRow = GetOriginalRowNumber(references, row);
                    if (originalRow > breakRow) {
                        if (!IsMergedCellContinuationRow(plan.ExportData.MergedCells, row, references.GetLength(1))) {
                            offsets.Add(row);
                        }

                        break;
                    }
                }
            }

            return offsets.ToList();
        }

        private static List<int> GetManualColumnBreakOffsets(WorksheetPdfExportPlan plan) {
            var offsets = new SortedSet<int>();
            string?[,]? references = plan.ExportData.CellReferences;
            if (references == null) {
                return offsets.ToList();
            }

            int rows = Math.Min(plan.ExportedRows, references.GetLength(0));
            int columns = references.GetLength(1);
            foreach (int breakColumn in plan.ManualColumnBreaks) {
                for (int column = 0; column < columns; column++) {
                    int originalColumn = GetOriginalColumnNumber(references, column, rows);
                    if (originalColumn > breakColumn) {
                        if (!IsMergedCellContinuationColumn(plan.ExportData.MergedCells, column, rows)) {
                            offsets.Add(column);
                        }

                        break;
                    }
                }
            }

            return offsets.ToList();
        }

        private static int GetOriginalRowNumber(string?[,] references, int row) {
            int columns = references.GetLength(1);
            for (int column = 0; column < columns; column++) {
                string? reference = references[row, column];
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
                if (cell.Row > 0) {
                    return cell.Row;
                }
            }

            return 0;
        }

        private static int GetOriginalColumnNumber(string?[,] references, int column, int rows) {
            for (int row = 0; row < rows; row++) {
                string? reference = references[row, column];
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                (int Row, int Col) cell = A1.ParseCellRef(reference!.Replace("$", string.Empty));
                if (cell.Col > 0) {
                    return cell.Col;
                }
            }

            return 0;
        }

        private static bool IsMergedCellContinuationRow(MergeLayoutData? mergedCells, int row, int columns) {
            if (mergedCells == null) {
                return false;
            }

            for (int column = 0; column < columns; column++) {
                if (mergedCells.IsContinuation(row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsMergedCellContinuationColumn(MergeLayoutData? mergedCells, int column, int rows) {
            if (mergedCells == null) {
                return false;
            }

            for (int row = 0; row < rows; row++) {
                if (mergedCells.IsContinuation(row, column)) {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<PdfCore.PdfTableCell[]> CreatePdfRows(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>>? imagesByCellReference, IReadOnlyList<int> rowIndexes, int startColumn, int columnCount, string emptyCellText, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, string sheetName) {
            int endColumn = Math.Min(values.GetLength(1), startColumn + columnCount);
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= values.GetLength(0)) {
                    continue;
                }

                var cells = new List<PdfCore.PdfTableCell>(columnCount);
                for (int column = startColumn; column < endColumn; column++) {
                    if (mergedCells?.IsContinuation(row, column) == true) {
                        continue;
                    }

                    ExcelCellStyleSnapshot? style = GetCellStyle(styles, row, column);
                    ExcelHyperlinkSnapshot? hyperlink = GetHyperlink(hyperlinks, row, column);
                    string text = FormatCellValue(values[row, column], style, emptyCellText);
                    MergeSpan? span = ClipMergeSpanToChunk(mergedCells?.GetSpan(row, column), row, rowIndexes, localRow, column, endColumn);
                    string? cellDestinationName = TryGetCellDestinationName(cellReferences, row, column, sheetName, cellDestinations, out string? destinationName)
                        ? destinationName
                        : null;
                    IReadOnlyList<WorksheetImageExportData>? cellImages = GetCellImages(imagesByCellReference, cellReferences, row, column);
                    cells.Add(CreatePdfCell(text, style, hyperlink, span, sheetDestinations, cellDestinations, sheetName, cellDestinationName, cellImages));
                }

                yield return cells.ToArray();
            }
        }

        private static MergeSpan? ClipMergeSpanToChunk(MergeSpan? span, int row, IReadOnlyList<int> rowIndexes, int localRow, int column, int endColumn) {
            if (span == null) {
                return span;
            }

            int clippedColumnSpan = Math.Min(span.ColumnSpan, Math.Max(1, endColumn - column));
            int contiguousRows = 1;
            for (int offset = 1; offset < span.RowSpan && localRow + offset < rowIndexes.Count; offset++) {
                if (rowIndexes[localRow + offset] != row + offset) {
                    break;
                }

                contiguousRows++;
            }

            int clippedRowSpan = Math.Min(span.RowSpan, contiguousRows);
            if (clippedRowSpan == span.RowSpan && clippedColumnSpan == span.ColumnSpan) {
                return span;
            }

            return new MergeSpan(clippedRowSpan, clippedColumnSpan);
        }

        private static IReadOnlyList<WorksheetImageExportData>? GetCellImages(IReadOnlyDictionary<string, IReadOnlyList<WorksheetImageExportData>>? imagesByCellReference, string?[,]? cellReferences, int row, int column) {
            if (imagesByCellReference == null || imagesByCellReference.Count == 0 || cellReferences == null || row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return null;
            }

            string? cellReference = cellReferences[row, column];
            if (string.IsNullOrWhiteSpace(cellReference)) {
                return null;
            }

            return imagesByCellReference.TryGetValue(NormalizeCellReference(cellReference!), out IReadOnlyList<WorksheetImageExportData>? images)
                ? images
                : null;
        }

        private static PdfCore.PdfTableCell CreatePdfCell(string text, ExcelCellStyleSnapshot? style, ExcelHyperlinkSnapshot? hyperlink, MergeSpan? span, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, string sheetName, string? cellDestinationName, IReadOnlyList<WorksheetImageExportData>? cellImages) {
            int rowSpan = span?.RowSpan ?? 1;
            int columnSpan = span?.ColumnSpan ?? 1;
            PdfCore.PdfColor? textColor = ToPdfColor(style?.FontColorHex);
            string? linkUri = hyperlink?.IsExternal == true ? hyperlink.Target : null;
            string? linkDestinationName = TryGetInternalHyperlinkDestinationName(hyperlink, sheetName, sheetDestinations, cellDestinations, out string? destinationName)
                ? destinationName
                : null;
            string? linkContents = linkUri == null && linkDestinationName == null ? null : text;
            IReadOnlyList<PdfCore.PdfTableCellImage>? pdfImages = ToPdfTableCellImages(cellImages);
            IReadOnlyList<PdfCore.TextRun> runs = style != null && (style.Bold || style.Italic || style.Underline || textColor.HasValue)
                ? new[] { new PdfCore.TextRun(text, bold: style.Bold, underline: style.Underline, color: textColor, italic: style.Italic) }
                : new[] { PdfCore.TextRun.Normal(text) };

            return new PdfCore.PdfTableCell(
                runs,
                columnSpan,
                linkUri,
                linkContents,
                rowSpan,
                images: pdfImages,
                linkDestinationName: linkDestinationName,
                namedDestinationName: cellDestinationName);
        }

        private static IReadOnlyList<PdfCore.PdfTableCellImage>? ToPdfTableCellImages(IReadOnlyList<WorksheetImageExportData>? images) {
            if (images == null || images.Count == 0) {
                return null;
            }

            var pdfImages = new List<PdfCore.PdfTableCellImage>(images.Count);
            foreach (WorksheetImageExportData image in images) {
                pdfImages.Add(new PdfCore.PdfTableCellImage(image.Bytes, image.WidthPoints, image.HeightPoints));
            }

            return pdfImages;
        }

        private static bool TryGetCellDestinationName(string?[,]? cellReferences, int row, int column, string sheetName, IReadOnlyDictionary<string, string> cellDestinations, out string? destinationName) {
            destinationName = null;
            if (cellReferences == null || row >= cellReferences.GetLength(0) || column >= cellReferences.GetLength(1)) {
                return false;
            }

            string? cellReference = cellReferences[row, column];
            return !string.IsNullOrWhiteSpace(cellReference) &&
                cellDestinations.TryGetValue(CreateCellDestinationKey(sheetName, cellReference!), out destinationName);
        }

        private static bool TryGetInternalHyperlinkDestinationName(ExcelHyperlinkSnapshot? hyperlink, string currentSheetName, IReadOnlyDictionary<string, string> sheetDestinations, IReadOnlyDictionary<string, string> cellDestinations, out string? destinationName) {
            destinationName = null;
            if (hyperlink == null || hyperlink.IsExternal || !TryParseInternalTarget(hyperlink.Target, currentSheetName, out string? sheetName, out string? cellReference)) {
                return false;
            }

            if (!string.IsNullOrEmpty(cellReference)) {
                return cellDestinations.TryGetValue(CreateCellDestinationKey(sheetName!, cellReference!), out destinationName);
            }

            return sheetDestinations.TryGetValue(sheetName!, out destinationName);
        }

        private static bool TryParseInternalSheetName(string? value, string currentSheetName, out string? sheetName) {
            if (TryParseInternalTarget(value, currentSheetName, out sheetName, out _)) {
                return true;
            }

            sheetName = null;
            return false;
        }

        private static bool TryParseInternalTarget(string? value, string currentSheetName, out string? sheetName, out string? cellReference) {
            sheetName = null;
            cellReference = null;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string trimmedValue = value!.Trim();
            int bangIndex = trimmedValue.LastIndexOf('!');
            if (bangIndex < 0) {
                string sameSheetReferenceToken = trimmedValue.Replace("$", string.Empty);
                if (!TryGetTopLeftCellReference(sameSheetReferenceToken, out string? sameSheetReference)) {
                    return false;
                }

                sheetName = currentSheetName;
                cellReference = sameSheetReference;
                return true;
            }

            if (bangIndex == 0 || bangIndex >= trimmedValue.Length - 1) {
                return false;
            }

            string sheetToken = trimmedValue.Substring(0, bangIndex).Trim();
            string referenceToken = trimmedValue.Substring(bangIndex + 1).Trim().Replace("$", string.Empty);
            if (sheetToken.Length == 0 || sheetToken.IndexOf('[') >= 0 || sheetToken.IndexOf(']') >= 0) {
                return false;
            }

            if (!TryGetTopLeftCellReference(referenceToken, out string? normalizedReference)) {
                return false;
            }

            string unquoted = UnquoteInternalSheetName(sheetToken);
            if (unquoted.Length == 0) {
                return false;
            }

            sheetName = unquoted;
            cellReference = normalizedReference;
            return true;
        }

        private static bool TryGetTopLeftCellReference(string referenceToken, out string? cellReference) {
            cellReference = null;
            if (string.IsNullOrWhiteSpace(referenceToken)) {
                return false;
            }

            string token = referenceToken.Trim();
            if (A1.TryParseRange(token, out int firstRow, out int firstColumn, out _, out _)) {
                cellReference = A1.CellReference(firstRow, firstColumn);
                return true;
            }

            (int Row, int Col) cell = A1.ParseCellRef(token);
            if (cell.Row <= 0 || cell.Col <= 0) {
                return false;
            }

            cellReference = A1.CellReference(cell.Row, cell.Col);
            return true;
        }

        private static string UnquoteInternalSheetName(string sheetToken) {
            string trimmedToken = sheetToken.Trim();
            if (trimmedToken.Length >= 2 && trimmedToken[0] == '\'' && trimmedToken[trimmedToken.Length - 1] == '\'') {
                return trimmedToken.Substring(1, trimmedToken.Length - 2).Replace("''", "'");
            }

            return trimmedToken;
        }

        private static PdfCore.PdfTableStyle CreateTableStyle(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup, IReadOnlyList<int> rowIndexes, int headerRowCount, ExcelCellStyleSnapshot?[,]? styles, ConditionalFillData? conditionalFills, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int columnOffset = 0, int exportedColumns = 0) {
            int exportedRows = rowIndexes.Count;
            int headerRows = Math.Min(headerRowCount, exportedRows);
            var tableStyle = new PdfCore.PdfTableStyle {
                HeaderRowCount = headerRows,
                RepeatHeaderRowCount = headerRows == 0 ? null : headerRows,
                CellPaddingX = 4,
                CellPaddingY = 3,
                HeaderFill = PdfCore.PdfColor.FromRgb(230, 238, 247),
                HeaderTextColor = PdfCore.PdfColor.FromRgb(31, 78, 121),
                RowStripeFill = PdfCore.PdfColor.FromRgb(248, 250, 252)
            };

            if (columnWidths != null) {
                List<double> widthWeights = columnWidths.WidthWeights.Skip(columnOffset).Take(exportedColumns).ToList();
                tableStyle.ColumnWidthWeights = widthWeights.Count == 0 ? columnWidths.WidthWeights : widthWeights;
                if (IsFitToWidth(pageSetup)) {
                    tableStyle.MaxWidth = CalculateFitToWidthMaxWidth(options, pageSetup);
                } else if (pageSetup?.Scale is uint scale && scale > 0U && scale < 100U) {
                    double approximateWidth = CalculateChunkApproximateWidth(columnWidths, columnOffset, exportedColumns);
                    tableStyle.MaxWidth = Math.Max(24D, approximateWidth * scale / 100D);
                }
            }

            if (rowHeights != null) {
                tableStyle.RowMinHeights = rowIndexes
                    .Select(row => row >= 0 && row < rowHeights.MinHeights.Count ? rowHeights.MinHeights[row] : null)
                    .ToList();
            }

            Dictionary<(int Row, int Column), PdfCore.PdfColor>? cellFills = CreateCellFills(styles, conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellFills != null) {
                tableStyle.CellFills = cellFills;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? cellDataBars = CreateCellDataBars(conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellDataBars != null) {
                tableStyle.CellDataBars = cellDataBars;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? cellIcons = CreateCellIcons(conditionalFills, rowIndexes, columnOffset, exportedColumns);
            if (cellIcons != null) {
                tableStyle.CellIcons = cellIcons;
                tableStyle.CellPaddings = CreateIconCellPaddings(cellIcons, tableStyle.CellPaddings);
            }

            Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? cellAlignments = CreateCellAlignments(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellAlignments != null) {
                tableStyle.CellAlignments = cellAlignments;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? cellVerticalAlignments = CreateCellVerticalAlignments(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellVerticalAlignments != null) {
                tableStyle.CellVerticalAlignments = cellVerticalAlignments;
            }

            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = CreateCellBorders(styles, rowIndexes, columnOffset, exportedColumns);
            if (cellBorders != null) {
                tableStyle.CellBorders = cellBorders;
            }

            return tableStyle;
        }

        private static double CalculateChunkApproximateWidth(ColumnLayoutData columnWidths, int columnOffset, int exportedColumns) {
            if (exportedColumns <= 0 || columnOffset <= 0 && exportedColumns >= columnWidths.WidthWeights.Count) {
                return columnWidths.ApproximateWidthPoints;
            }

            double totalWeight = columnWidths.WidthWeights.Sum();
            double chunkWeight = columnWidths.WidthWeights.Skip(columnOffset).Take(exportedColumns).Sum();
            if (totalWeight <= 0D || chunkWeight <= 0D) {
                return columnWidths.ApproximateWidthPoints;
            }

            return columnWidths.ApproximateWidthPoints * chunkWeight / totalWeight;
        }

        private static bool IsFitToWidth(ExcelSheetPageSetup? pageSetup) {
            return pageSetup?.FitToWidth is uint fitToWidth && fitToWidth > 0U;
        }

        private static double CalculateFitToWidthMaxWidth(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            PdfCore.PageSize pageSize = GetEffectivePageSize(options, pageSetup);
            PdfCore.PageMargins margins = GetEffectiveMargins(options, pageSetup);
            return Math.Max(24D, pageSize.Width - margins.Left - margins.Right);
        }

        private static PdfCore.PageSize GetEffectivePageSize(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            PdfCore.PageSize pageSize = options.PageSize ?? PdfCore.PageSizes.Letter;
            if (pageSetup?.Orientation == ExcelPageOrientation.Landscape) {
                return pageSize.Landscape();
            }

            if (pageSetup?.Orientation == ExcelPageOrientation.Portrait) {
                return pageSize.Portrait();
            }

            return pageSize;
        }

        private static PdfCore.PageMargins GetEffectiveMargins(ExcelPdfSaveOptions options, ExcelSheetPageSetup? pageSetup) {
            if (options.Margins.HasValue) {
                return options.Margins.Value;
            }

            if (pageSetup?.Margins != null) {
                return ToPdfMargins(pageSetup.Margins);
            }

            return PdfCore.PageMargins.Normal;
        }

        private static ExcelCellStyleSnapshot? GetCellStyle(ExcelCellStyleSnapshot?[,]? styles, int row, int column) {
            if (styles == null || row >= styles.GetLength(0) || column >= styles.GetLength(1)) {
                return null;
            }

            return styles[row, column];
        }

        private static ExcelHyperlinkSnapshot? GetHyperlink(ExcelHyperlinkSnapshot?[,]? hyperlinks, int row, int column) {
            if (hyperlinks == null || row >= hyperlinks.GetLength(0) || column >= hyperlinks.GetLength(1)) {
                return null;
            }

            return hyperlinks[row, column];
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfColor>? CreateCellFills(ExcelCellStyleSnapshot?[,]? styles, ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null && conditionalFills == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfColor>? fills = null;
            if (styles != null) {
                int columns = Math.Min(columnEnd, styles.GetLength(1));
                for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                    int row = rowIndexes[localRow];
                    if (row < 0 || row >= styles.GetLength(0)) {
                        continue;
                    }

                    for (int column = columnOffset; column < columns; column++) {
                        PdfCore.PdfColor? fill = ToPdfColor(styles[row, column]?.FillColorHex);
                        if (fill.HasValue) {
                            fills ??= new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
                            fills[(localRow, column - columnOffset)] = fill.Value;
                        }
                    }
                }
            }

            if (conditionalFills != null) {
                foreach (KeyValuePair<(int Row, int Column), string> conditionalFill in conditionalFills.FillColors) {
                    int localRow = FindLocalRowIndex(rowIndexes, conditionalFill.Key.Row);
                    if (localRow < 0 ||
                        conditionalFill.Key.Column < columnOffset ||
                        conditionalFill.Key.Column >= columnEnd) {
                        continue;
                    }

                    PdfCore.PdfColor? fill = ToPdfColor(conditionalFill.Value);
                    if (fill.HasValue) {
                        fills ??= new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
                        fills[(localRow, conditionalFill.Key.Column - columnOffset)] = fill.Value;
                    }
                }
            }

            return fills;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? CreateCellDataBars(ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (conditionalFills == null || conditionalFills.DataBars.Count == 0) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>? dataBars = null;
            foreach (KeyValuePair<(int Row, int Column), ConditionalDataBarCell> conditionalDataBar in conditionalFills.DataBars) {
                int localRow = FindLocalRowIndex(rowIndexes, conditionalDataBar.Key.Row);
                if (localRow < 0 ||
                    conditionalDataBar.Key.Column < columnOffset ||
                    conditionalDataBar.Key.Column >= columnEnd) {
                    continue;
                }

                PdfCore.PdfColor? fill = ToPdfColor(conditionalDataBar.Value.Color);
                if (fill.HasValue) {
                    dataBars ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellDataBar>();
                    dataBars[(localRow, conditionalDataBar.Key.Column - columnOffset)] = new PdfCore.PdfCellDataBar {
                        Color = fill.Value,
                        Ratio = conditionalDataBar.Value.Ratio
                    };
                }
            }

            return dataBars;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? CreateCellIcons(ConditionalFillData? conditionalFills, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (conditionalFills == null || conditionalFills.Icons.Count == 0) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : int.MaxValue;
            Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>? icons = null;
            foreach (KeyValuePair<(int Row, int Column), ConditionalIconCell> conditionalIcon in conditionalFills.Icons) {
                int localRow = FindLocalRowIndex(rowIndexes, conditionalIcon.Key.Row);
                if (localRow < 0 ||
                    conditionalIcon.Key.Column < columnOffset ||
                    conditionalIcon.Key.Column >= columnEnd) {
                    continue;
                }

                icons ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellIcon>();
                icons[(localRow, conditionalIcon.Key.Column - columnOffset)] = new PdfCore.PdfCellIcon {
                    Kind = conditionalIcon.Value.Kind,
                    Color = conditionalIcon.Value.Color,
                    Size = 8D
                };
            }

            return icons;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellPadding> CreateIconCellPaddings(IReadOnlyDictionary<(int Row, int Column), PdfCore.PdfCellIcon> icons, Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>? existingPaddings) {
            var paddings = existingPaddings == null
                ? new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>()
                : new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>(existingPaddings);

            foreach (KeyValuePair<(int Row, int Column), PdfCore.PdfCellIcon> icon in icons) {
                if (!paddings.TryGetValue(icon.Key, out PdfCore.PdfCellPadding? padding)) {
                    padding = new PdfCore.PdfCellPadding();
                } else {
                    padding = padding.Clone();
                }

                double requiredLeftPadding = icon.Value.Size + 8D;
                padding.Left = Math.Max(padding.Left ?? 0D, requiredLeftPadding);
                paddings[icon.Key] = padding;
            }

            return paddings;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? CreateCellAlignments(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? alignments = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfColumnAlign? alignment = ToPdfHorizontalAlignment(styles[row, column]?.HorizontalAlignment);
                    if (alignment.HasValue) {
                        alignments ??= new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
                        alignments[(localRow, column - columnOffset)] = alignment.Value;
                    }
                }
            }

            return alignments;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? CreateCellVerticalAlignments(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? alignments = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfCellVerticalAlign? alignment = ToPdfVerticalAlignment(styles[row, column]?.VerticalAlignment);
                    if (alignment.HasValue) {
                        alignments ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
                        alignments[(localRow, column - columnOffset)] = alignment.Value;
                    }
                }
            }

            return alignments;
        }

        private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateCellBorders(ExcelCellStyleSnapshot?[,]? styles, IReadOnlyList<int> rowIndexes, int columnOffset = 0, int exportedColumns = 0) {
            if (styles == null) {
                return null;
            }

            int columnEnd = exportedColumns > 0 ? columnOffset + exportedColumns : styles.GetLength(1);
            int columns = Math.Min(columnEnd, styles.GetLength(1));
            Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? borders = null;
            for (int localRow = 0; localRow < rowIndexes.Count; localRow++) {
                int row = rowIndexes[localRow];
                if (row < 0 || row >= styles.GetLength(0)) {
                    continue;
                }

                for (int column = columnOffset; column < columns; column++) {
                    PdfCore.PdfCellBorder? border = ToPdfCellBorder(styles[row, column]?.Border);
                    if (border != null) {
                        borders ??= new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
                        borders[(localRow, column - columnOffset)] = border;
                    }
                }
            }

            return borders;
        }

        private static int FindLocalRowIndex(IReadOnlyList<int> rowIndexes, int row) {
            for (int index = 0; index < rowIndexes.Count; index++) {
                if (rowIndexes[index] == row) {
                    return index;
                }
            }

            return -1;
        }

        private static PdfCore.PdfColumnAlign? ToPdfHorizontalAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToLowerInvariant()) {
                case "left":
                    return PdfCore.PdfColumnAlign.Left;
                case "center":
                case "centercontinuous":
                    return PdfCore.PdfColumnAlign.Center;
                case "right":
                    return PdfCore.PdfColumnAlign.Right;
                default:
                    return null;
            }
        }

        private static PdfCore.PdfCellVerticalAlign? ToPdfVerticalAlignment(string? alignment) {
            if (string.IsNullOrWhiteSpace(alignment)) {
                return null;
            }

            switch (alignment!.Trim().ToLowerInvariant()) {
                case "top":
                    return PdfCore.PdfCellVerticalAlign.Top;
                case "center":
                    return PdfCore.PdfCellVerticalAlign.Middle;
                case "bottom":
                    return PdfCore.PdfCellVerticalAlign.Bottom;
                default:
                    return null;
            }
        }

        private static PdfCore.PdfCellBorder? ToPdfCellBorder(ExcelCellBorderSnapshot? border) {
            if (border == null) {
                return null;
            }

            PdfCore.PdfCellBorderSide? left = ToPdfCellBorderSide(border.Left);
            PdfCore.PdfCellBorderSide? right = ToPdfCellBorderSide(border.Right);
            PdfCore.PdfCellBorderSide? top = ToPdfCellBorderSide(border.Top);
            PdfCore.PdfCellBorderSide? bottom = ToPdfCellBorderSide(border.Bottom);
            PdfCore.PdfCellBorderSide? diagonal = ToPdfCellBorderSide(border.Diagonal);
            bool hasDiagonalUp = border.DiagonalUp && diagonal != null;
            bool hasDiagonalDown = border.DiagonalDown && diagonal != null;
            if (left == null && right == null && top == null && bottom == null && !hasDiagonalUp && !hasDiagonalDown) {
                return null;
            }

            return new PdfCore.PdfCellBorder {
                Color = null,
                TopBorder = top,
                RightBorder = right,
                BottomBorder = bottom,
                LeftBorder = left,
                DiagonalUp = hasDiagonalUp,
                DiagonalDown = hasDiagonalDown,
                DiagonalUpBorder = hasDiagonalUp ? diagonal : null,
                DiagonalDownBorder = hasDiagonalDown ? diagonal : null
            };
        }

        private static PdfCore.PdfCellBorderSide? ToPdfCellBorderSide(ExcelBorderSideSnapshot? side) {
            if (side == null) {
                return null;
            }

            double width = ToPdfBorderWidth(side.Style);
            if (width <= 0) {
                return null;
            }

            return new PdfCore.PdfCellBorderSide {
                Color = ToPdfColor(side.ColorArgb) ?? PdfCore.PdfColor.FromRgb(0, 0, 0),
                Width = width,
                DashStyle = ToPdfBorderDashStyle(side.Style),
                LineStyle = ToPdfBorderLineStyle(side.Style)
            };
        }

        private static double ToPdfBorderWidth(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return 0D;
            }

            switch (style!.Trim().ToLowerInvariant()) {
                case "none":
                    return 0D;
                case "hair":
                    return 0.25D;
                case "medium":
                case "mediumdashdot":
                case "mediumdashdotdot":
                case "mediumdashed":
                    return 1.25D;
                case "thick":
                case "double":
                    return 2D;
                default:
                    return 0.5D;
            }
        }

        private static OfficeIMO.Drawing.OfficeStrokeDashStyle ToPdfBorderDashStyle(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            }

            switch (style!.Trim().ToLowerInvariant()) {
                case "dashed":
                case "mediumdashed":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dash;
                case "dotted":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Dot;
                case "dashdot":
                case "dashdotdot":
                case "mediumdashdot":
                case "mediumdashdotdot":
                case "slantdashdot":
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.DashDot;
                default:
                    return OfficeIMO.Drawing.OfficeStrokeDashStyle.Solid;
            }
        }

        private static PdfCore.PdfCellBorderLineStyle ToPdfBorderLineStyle(string? style) {
            string? normalized = style?.Trim();
            return !string.IsNullOrWhiteSpace(normalized) &&
                   string.Equals(normalized, "double", StringComparison.OrdinalIgnoreCase)
                ? PdfCore.PdfCellBorderLineStyle.TwoLine
                : PdfCore.PdfCellBorderLineStyle.Standard;
        }

        private static PdfCore.PdfColor? ToPdfColor(string? hex) {
            if (string.IsNullOrWhiteSpace(hex)) {
                return null;
            }

            string value = hex!.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }

            if (value.Length == 8) {
                value = value.Substring(2);
            }

            if (value.Length != 6 ||
                !byte.TryParse(value.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte r) ||
                !byte.TryParse(value.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte g) ||
                !byte.TryParse(value.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte b)) {
                return null;
            }

            return PdfCore.PdfColor.FromRgb(r, g, b);
        }

        private static string FormatCellValue(object? value, ExcelCellStyleSnapshot? style, string emptyCellText) {
            if (value == null) {
                return emptyCellText;
            }

            string? formatCode = style?.NumberFormatCode;
            if (!string.IsNullOrWhiteSpace(formatCode)) {
                string? formatted = TryFormatCellValue(value, style!, formatCode!);
                if (formatted != null) {
                    return formatted;
                }
            }

            if (value is IFormattable formattable) {
                return formattable.ToString(null, CultureInfo.InvariantCulture) ?? emptyCellText;
            }

            return value.ToString() ?? emptyCellText;
        }

        private static string? TryFormatCellValue(object value, ExcelCellStyleSnapshot style, string formatCode) {
            string normalized = GetNumberFormatSection(formatCode, 0).Trim();
            if (normalized.Length == 0 ||
                string.Equals(normalized, "General", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "@", StringComparison.Ordinal)) {
                return null;
            }

            if (value is DateTime dateValue || style.IsDateLike) {
                DateTime date = value is DateTime directDate
                    ? directDate
                    : TryGetDouble(value, out double oaDate) ? DateTime.FromOADate(oaDate) : default;
                if (date != default) {
                    return date.ToString(ToDotNetDateTimeFormat(normalized), CultureInfo.InvariantCulture);
                }
            }

            if (!TryGetDouble(value, out double number)) {
                return null;
            }

            normalized = GetNumberFormatSection(formatCode, GetNumberFormatSectionIndex(number)).Trim();
            if (normalized.Length == 0 ||
                string.Equals(normalized, "General", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "@", StringComparison.Ordinal)) {
                return null;
            }

            if (normalized.IndexOf('%') >= 0) {
                int decimals = CountDecimalPlaces(normalized);
                bool wrapPercent = ShouldWrapNegativeNumber(normalized, number);
                double percentNumber = wrapPercent ? Math.Abs(number) : number;
                string numeric = (percentNumber * 100D).ToString(decimals > 0 ? "N" + decimals.ToString(CultureInfo.InvariantCulture) : "N0", CultureInfo.InvariantCulture);
                if (wrapPercent) {
                    return "(" + numeric + "%)";
                }

                return numeric + "%";
            }

            string? prefix = ExtractQuotedLiteralPrefix(normalized);
            bool useGrouping = normalized.IndexOf(',') >= 0;
            int decimalPlaces = CountDecimalPlaces(normalized);
            string numberFormat = (useGrouping ? "N" : "F") + decimalPlaces.ToString(CultureInfo.InvariantCulture);
            bool wrapNumber = ShouldWrapNegativeNumber(normalized, number);
            double displayNumber = wrapNumber ? Math.Abs(number) : number;
            string numericValue = (prefix ?? string.Empty) + displayNumber.ToString(numberFormat, CultureInfo.InvariantCulture);
            return wrapNumber ? "(" + numericValue + ")" : numericValue;
        }

        private static bool TryGetDouble(object value, out double number) {
            switch (value) {
                case double doubleValue:
                    number = doubleValue;
                    return true;
                case float floatValue:
                    number = floatValue;
                    return true;
                case decimal decimalValue:
                    number = (double)decimalValue;
                    return true;
                case int intValue:
                    number = intValue;
                    return true;
                case long longValue:
                    number = longValue;
                    return true;
                case short shortValue:
                    number = shortValue;
                    return true;
                case byte byteValue:
                    number = byteValue;
                    return true;
                default:
                    return double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number);
            }
        }

        private static string GetNumberFormatSection(string formatCode, int sectionIndex) {
            string[] sections = formatCode.Split(';');
            if (sectionIndex >= 0 && sectionIndex < sections.Length) {
                return sections[sectionIndex];
            }

            return sections.Length > 0 ? sections[0] : formatCode;
        }

        private static int GetNumberFormatSectionIndex(double number) {
            if (number < 0D) {
                return 1;
            }

            return number == 0D ? 2 : 0;
        }

        private static bool ShouldWrapNegativeNumber(string formatCode, double value) =>
            value < 0D && formatCode.IndexOf('(') >= 0 && formatCode.IndexOf(')') > formatCode.IndexOf('(');

        private static int CountDecimalPlaces(string formatCode) {
            int decimalIndex = formatCode.IndexOf('.');
            if (decimalIndex < 0) {
                return 0;
            }

            int count = 0;
            for (int i = decimalIndex + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (ch == '0' || ch == '#') {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static string? ExtractQuotedLiteralPrefix(string formatCode) {
            int quoteStart = formatCode.IndexOf('"');
            if (quoteStart < 0) {
                return null;
            }

            int quoteEnd = formatCode.IndexOf('"', quoteStart + 1);
            if (quoteEnd <= quoteStart + 1) {
                return null;
            }

            string literal = formatCode.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
            return literal.Length == 0 ? null : literal;
        }

        private static string ToDotNetDateTimeFormat(string excelFormat) {
            string format = StripExcelBracketAndColorTokens(excelFormat);
            format = ReplaceExcelDateTokens(format);
            format = ReplaceIgnoreCase(format, "AM/PM", "tt");
            format = ReplaceIgnoreCase(format, "A/P", "tt");
            return format;
        }

        private static string ReplaceIgnoreCase(string value, string oldValue, string newValue) {
            return value
                .Replace(oldValue, newValue)
                .Replace(oldValue.ToLowerInvariant(), newValue)
                .Replace(oldValue.ToUpperInvariant(), newValue);
        }

        private static string StripExcelBracketAndColorTokens(string format) {
            var builder = new System.Text.StringBuilder(format.Length);
            for (int i = 0; i < format.Length; i++) {
                char ch = format[i];
                if (ch == '[') {
                    int close = format.IndexOf(']', i + 1);
                    if (close >= 0) {
                        string token = format.Substring(i + 1, close - i - 1);
                        if (token.All(c => c == 'h' || c == 'H' || c == 'm' || c == 'M' || c == 's' || c == 'S')) {
                            builder.Append(token);
                        }

                        i = close;
                        continue;
                    }
                }

                if (ch == '\\' || ch == '_') {
                    if (i + 1 < format.Length) {
                        builder.Append(format[i + 1]);
                        i++;
                    }
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static string ReplaceExcelDateTokens(string format) {
            var builder = new System.Text.StringBuilder(format.Length);
            for (int i = 0; i < format.Length;) {
                char ch = format[i];
                if (ch == '"') {
                    int end = format.IndexOf('"', i + 1);
                    if (end < 0) {
                        break;
                    }

                    builder.Append('\'').Append(format.Substring(i + 1, end - i - 1).Replace("'", "\\'")).Append('\'');
                    i = end + 1;
                    continue;
                }

                if (!IsExcelDateFormatLetter(ch)) {
                    builder.Append(ch);
                    i++;
                    continue;
                }

                int start = i;
                while (i < format.Length && char.ToLowerInvariant(format[i]) == char.ToLowerInvariant(ch)) {
                    i++;
                }

                string token = format.Substring(start, i - start);
                builder.Append(ConvertExcelDateToken(token, builder, format, i));
            }

            return builder.ToString();
        }

        private static bool IsExcelDateFormatLetter(char ch) {
            switch (char.ToLowerInvariant(ch)) {
                case 'y':
                case 'm':
                case 'd':
                case 'h':
                case 's':
                    return true;
                default:
                    return false;
            }
        }

        private static string ConvertExcelDateToken(string token, System.Text.StringBuilder output, string format, int nextIndex) {
            char lower = char.ToLowerInvariant(token[0]);
            switch (lower) {
                case 'y':
                    return token.Length <= 2 ? "yy" : "yyyy";
                case 'd':
                    return token.Length <= 1 ? "d" : token.Length == 2 ? "dd" : token.Length == 3 ? "ddd" : "dddd";
                case 'h':
                    return token.Length <= 1 ? "h" : "hh";
                case 's':
                    return token.Length <= 1 ? "s" : "ss";
                case 'm':
                    bool timeMinute = PreviousNonSpace(output) == ':' || NextNonSpace(format, nextIndex) == ':';
                    if (timeMinute) {
                        return token.Length <= 1 ? "m" : "mm";
                    }

                    return token.Length <= 1 ? "M" : token.Length == 2 ? "MM" : token.Length == 3 ? "MMM" : "MMMM";
                default:
                    return token;
            }
        }

        private static char PreviousNonSpace(System.Text.StringBuilder builder) {
            for (int i = builder.Length - 1; i >= 0; i--) {
                if (!char.IsWhiteSpace(builder[i])) {
                    return builder[i];
                }
            }

            return '\0';
        }

        private static char NextNonSpace(string value, int startIndex) {
            for (int i = startIndex; i < value.Length; i++) {
                if (!char.IsWhiteSpace(value[i])) {
                    return value[i];
                }
            }

            return '\0';
        }

        private sealed class WorksheetPdfExportPlan {
            public WorksheetPdfExportPlan(string sheetName, ExcelSheetPageSetup? pageSetup, ExcelSheet.HeaderFooterSnapshot? headerFooter, SheetExportData exportData, IReadOnlyList<WorksheetImageExportData> images, IReadOnlyList<WorksheetChartExportData> charts, bool hasTable, int exportedRows, IReadOnlyList<int> manualRowBreaks, IReadOnlyList<int> manualColumnBreaks, string bookmarkName) {
                SheetName = sheetName;
                PageSetup = pageSetup;
                HeaderFooter = headerFooter;
                ExportData = exportData;
                Images = images;
                Charts = charts;
                HasTable = hasTable;
                ExportedRows = exportedRows;
                ManualRowBreaks = manualRowBreaks;
                ManualColumnBreaks = manualColumnBreaks;
                BookmarkName = bookmarkName;
            }

            public string SheetName { get; }
            public ExcelSheetPageSetup? PageSetup { get; }
            public ExcelSheet.HeaderFooterSnapshot? HeaderFooter { get; }
            public SheetExportData ExportData { get; }
            public IReadOnlyList<WorksheetImageExportData> Images { get; }
            public IReadOnlyList<WorksheetChartExportData> Charts { get; }
            public bool HasTable { get; }
            public int ExportedRows { get; }
            public IReadOnlyList<int> ManualRowBreaks { get; }
            public IReadOnlyList<int> ManualColumnBreaks { get; }
            public string BookmarkName { get; }
        }

        private sealed class SheetExportData {
            public SheetExportData(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights, int headerRowCount, int firstBodyRowNumber, ConditionalFillData? conditionalFills = null) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
                HeaderRowCount = headerRowCount;
                FirstBodyRowNumber = firstBodyRowNumber;
                ConditionalFills = conditionalFills;
            }

            public object?[,] Values { get; }
            public ExcelCellStyleSnapshot?[,]? Styles { get; }
            public ExcelHyperlinkSnapshot?[,]? Hyperlinks { get; }
            public string?[,]? CellReferences { get; }
            public MergeLayoutData? MergedCells { get; }
            public ColumnLayoutData? ColumnWidths { get; }
            public RowLayoutData? RowHeights { get; }
            public int HeaderRowCount { get; }
            public int FirstBodyRowNumber { get; }
            public ConditionalFillData? ConditionalFills { get; }
        }

        private sealed class ConditionalFillData {
            public ConditionalFillData(IReadOnlyDictionary<(int Row, int Column), string> fillColors, IReadOnlyDictionary<(int Row, int Column), ConditionalDataBarCell> dataBars, IReadOnlyDictionary<(int Row, int Column), ConditionalIconCell> icons) {
                FillColors = fillColors;
                DataBars = dataBars;
                Icons = icons;
            }

            public IReadOnlyDictionary<(int Row, int Column), string> FillColors { get; }
            public IReadOnlyDictionary<(int Row, int Column), ConditionalDataBarCell> DataBars { get; }
            public IReadOnlyDictionary<(int Row, int Column), ConditionalIconCell> Icons { get; }
        }

        private sealed class ConditionalDataBarCell {
            public ConditionalDataBarCell(string color, double ratio) {
                Color = color;
                Ratio = ratio;
            }

            public string Color { get; }
            public double Ratio { get; }
        }

        private sealed class ConditionalIconCell {
            public ConditionalIconCell(PdfCore.PdfCellIconKind kind, PdfCore.PdfColor color) {
                Kind = kind;
                Color = color;
            }

            public PdfCore.PdfCellIconKind Kind { get; }
            public PdfCore.PdfColor Color { get; }
        }

        private sealed class WorksheetImageExportData {
            public WorksheetImageExportData(byte[] bytes, double widthPoints, double heightPoints, string cellReference) {
                Bytes = bytes;
                WidthPoints = widthPoints;
                HeightPoints = heightPoints;
                CellReference = cellReference;
            }

            public byte[] Bytes { get; }
            public double WidthPoints { get; }
            public double HeightPoints { get; }
            public string CellReference { get; }
        }

        private sealed class WorksheetChartExportData {
            public WorksheetChartExportData(ExcelChartSnapshot snapshot) {
                Snapshot = snapshot;
            }

            public ExcelChartSnapshot Snapshot { get; }
        }

        private sealed class RangeExportData {
            public RangeExportData(object?[,] values, ExcelCellStyleSnapshot?[,]? styles, ExcelHyperlinkSnapshot?[,]? hyperlinks, string?[,]? cellReferences, MergeLayoutData? mergedCells, ColumnLayoutData? columnWidths, RowLayoutData? rowHeights) {
                Values = values;
                Styles = styles;
                Hyperlinks = hyperlinks;
                CellReferences = cellReferences;
                MergedCells = mergedCells;
                ColumnWidths = columnWidths;
                RowHeights = rowHeights;
            }

            public object?[,] Values { get; }
            public ExcelCellStyleSnapshot?[,]? Styles { get; }
            public ExcelHyperlinkSnapshot?[,]? Hyperlinks { get; }
            public string?[,]? CellReferences { get; }
            public MergeLayoutData? MergedCells { get; }
            public ColumnLayoutData? ColumnWidths { get; }
            public RowLayoutData? RowHeights { get; }
        }

        private sealed class VisibilityLayoutData {
            public VisibilityLayoutData(List<int> rowOffsets, List<int> columnOffsets, int originalRowCount, int originalColumnCount) {
                RowOffsets = rowOffsets;
                ColumnOffsets = columnOffsets;
                OriginalRowCount = originalRowCount;
                OriginalColumnCount = originalColumnCount;
            }

            public List<int> RowOffsets { get; }
            public List<int> ColumnOffsets { get; }
            public int OriginalRowCount { get; }
            public int OriginalColumnCount { get; }
        }

        private sealed class ColumnLayoutData {
            public ColumnLayoutData(List<double> widthWeights, double approximateWidthPoints) {
                WidthWeights = widthWeights;
                ApproximateWidthPoints = approximateWidthPoints;
            }

            public List<double> WidthWeights { get; }
            public double ApproximateWidthPoints { get; }
        }

        private sealed class RowLayoutData {
            public RowLayoutData(List<double?> minHeights) {
                MinHeights = minHeights;
            }

            public List<double?> MinHeights { get; }
        }

        private sealed class TableAxisChunk {
            public TableAxisChunk(int start, int count) {
                Start = start;
                Count = count;
            }

            public int Start { get; }
            public int Count { get; }
        }

        private sealed class TableChunk {
            public TableChunk(IReadOnlyList<int> rowIndexes, int headerRowCount, int startColumn, int columnCount) {
                RowIndexes = rowIndexes;
                HeaderRowCount = headerRowCount;
                StartColumn = startColumn;
                ColumnCount = columnCount;
            }

            public IReadOnlyList<int> RowIndexes { get; }
            public int HeaderRowCount { get; }
            public int StartColumn { get; }
            public int ColumnCount { get; }
        }

        private sealed class MergeLayoutData {
            private readonly MergeSpan?[,] _spans;
            private readonly bool[,] _continuations;

            public MergeLayoutData(int rowCount, int columnCount) {
                _spans = new MergeSpan?[rowCount, columnCount];
                _continuations = new bool[rowCount, columnCount];
            }

            public bool HasAny { get; private set; }

            public void SetSpan(int row, int column, int rowSpan, int columnSpan) {
                if (row < 0 || column < 0 || row >= _spans.GetLength(0) || column >= _spans.GetLength(1)) {
                    return;
                }

                rowSpan = Math.Min(rowSpan, _spans.GetLength(0) - row);
                columnSpan = Math.Min(columnSpan, _spans.GetLength(1) - column);
                if (rowSpan <= 1 && columnSpan <= 1) {
                    return;
                }

                _spans[row, column] = new MergeSpan(rowSpan, columnSpan);
                for (int r = row; r < row + rowSpan; r++) {
                    for (int c = column; c < column + columnSpan; c++) {
                        if (r != row || c != column) {
                            _continuations[r, c] = true;
                        }
                    }
                }

                HasAny = true;
            }

            public MergeSpan? GetSpan(int row, int column) =>
                row >= 0 && column >= 0 && row < _spans.GetLength(0) && column < _spans.GetLength(1)
                    ? _spans[row, column]
                    : null;

            public bool IsContinuation(int row, int column) =>
                row >= 0 && column >= 0 && row < _continuations.GetLength(0) && column < _continuations.GetLength(1) && _continuations[row, column];

            public void CopyTo(MergeLayoutData target, int rowOffset) {
                for (int row = 0; row < _spans.GetLength(0); row++) {
                    for (int column = 0; column < _spans.GetLength(1); column++) {
                        MergeSpan? span = _spans[row, column];
                        if (span != null) {
                            target.SetSpan(row + rowOffset, column, span.RowSpan, span.ColumnSpan);
                        }
                    }
                }
            }
        }

        private sealed class MergeSpan {
            public MergeSpan(int rowSpan, int columnSpan) {
                RowSpan = rowSpan;
                ColumnSpan = columnSpan;
            }

            public int RowSpan { get; }
            public int ColumnSpan { get; }
        }
    }
}
