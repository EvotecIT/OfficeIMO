using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private const string ExternalLinkPathRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        internal static ExcelDocument ToExcelDocument(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            ExcelDocument document = ExcelDocument.Create();
            if (workbook.Worksheets.Count == 0 && workbook.ChartSheets.Count == 0) {
                document.AddWorksheet("Sheet1");
            }

            foreach (LegacyXlsSheetProjectionEntry sheetEntry in EnumerateSheetsInWorkbookOrder(workbook)) {
                if (sheetEntry.Worksheet != null) {
                    ExcelSheet sheet = document.AddWorksheet(sheetEntry.Worksheet.Name);
                    ProjectWorksheet(workbook, sheetEntry.Worksheet, sheet);
                } else if (sheetEntry.ChartSheet != null) {
                    ProjectChartSheet(sheetEntry.ChartSheet, document);
                }
            }

            ProjectDefinedNames(workbook, document);
            ProjectAutoFilters(workbook, document);
            ProjectExternalReferences(workbook, document);
            ProjectExternalQueryConnections(workbook, document);
            ProjectDocumentProperties(workbook, document);
            ProjectWorkbookOptions(workbook, document);
            ProjectCodeNames(workbook, document);
            ProjectSheetTabIds(workbook, document);
            ProjectCalculationSettings(workbook, document);
            ProjectWorkbookWindow(workbook, document);
            ProjectWriteReservation(workbook, document);
            ProjectTableStyles(workbook, document);
            ProjectCellStyles(workbook, document);
            ProjectWorkbookTheme(workbook, document);

            if (workbook.Protection?.IsProtected == true || workbook.WindowsLocked.HasValue) {
                document.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
                    ProtectStructure = workbook.Protection?.IsProtected == true,
                    ProtectWindows = workbook.WindowsLocked == true,
                    LegacyPasswordHash = workbook.Protection?.LegacyPasswordHash
                });
            } else if (!string.IsNullOrWhiteSpace(workbook.Protection?.LegacyPasswordHash)) {
                WorkbookProtection protection = GetOrCreateWorkbookProtection(document);
                protection.WorkbookPassword = workbook.Protection!.LegacyPasswordHash;
            }

            ProjectWorkbookRevisionProtection(workbook, document);

            return document;
        }

        private static IReadOnlyList<LegacyXlsSheetProjectionEntry> EnumerateSheetsInWorkbookOrder(LegacyXlsWorkbook workbook) {
            var sheets = new List<LegacyXlsSheetProjectionEntry>(workbook.Worksheets.Count + workbook.ChartSheets.Count);
            foreach (LegacyXlsWorksheet worksheet in workbook.Worksheets) {
                sheets.Add(new LegacyXlsSheetProjectionEntry(worksheet.StreamOffset, worksheet, null));
            }

            foreach (LegacyXlsChartSheet chartSheet in workbook.ChartSheets) {
                sheets.Add(new LegacyXlsSheetProjectionEntry(chartSheet.StreamOffset, null, chartSheet));
            }

            sheets.Sort((left, right) => left.StreamOffset.CompareTo(right.StreamOffset));
            return sheets;
        }

        private readonly struct LegacyXlsSheetProjectionEntry {
            internal LegacyXlsSheetProjectionEntry(int streamOffset, LegacyXlsWorksheet? worksheet, LegacyXlsChartSheet? chartSheet) {
                StreamOffset = streamOffset;
                Worksheet = worksheet;
                ChartSheet = chartSheet;
            }

            internal int StreamOffset { get; }

            internal LegacyXlsWorksheet? Worksheet { get; }

            internal LegacyXlsChartSheet? ChartSheet { get; }
        }

        private static void ProjectWorkbookTheme(LegacyXlsWorkbook workbook, ExcelDocument document) {
            foreach (LegacyXlsThemeRecord themeRecord in workbook.ThemeRecords.Reverse()) {
                if (LegacyXlsThemePackageReader.TryExtractThemeXml(themeRecord, out string? themeXml)
                    && !string.IsNullOrWhiteSpace(themeXml)) {
                    document.SetWorkbookThemeXml(themeXml!);
                    return;
                }
            }

            if (workbook.ThemeRecords.Any(themeRecord => themeRecord.IsDefaultThemeMarker)) {
                document.EnsureWorkbookThemeAndStyles();
            }
        }

        private static void ProjectTableStyles(LegacyXlsWorkbook workbook, ExcelDocument document) {
            LegacyXlsTableStyleCollection? collection = workbook.TableStyleCollections.LastOrDefault();
            bool hasCustomStyles = workbook.TableStyles.Count > 0;
            bool hasDefaultStyleNames = collection != null
                && (!string.IsNullOrWhiteSpace(collection.DefaultTableStyleName)
                    || !string.IsNullOrWhiteSpace(collection.DefaultPivotStyleName));
            if (!hasCustomStyles && !hasDefaultStyleNames) {
                return;
            }

            document.EnsureWorkbookThemeAndStyles();
            Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            TableStyles tableStyles = stylesheet.TableStyles ??= new TableStyles();
            string? defaultTableStyleName = collection?.DefaultTableStyleName;
            if (!string.IsNullOrWhiteSpace(defaultTableStyleName)) {
                tableStyles.DefaultTableStyle = defaultTableStyleName;
            }

            string? defaultPivotStyleName = collection?.DefaultPivotStyleName;
            if (!string.IsNullOrWhiteSpace(defaultPivotStyleName)) {
                tableStyles.DefaultPivotStyle = defaultPivotStyleName;
            }

            tableStyles.RemoveAllChildren<DocumentFormat.OpenXml.Spreadsheet.TableStyle>();
            foreach (LegacyXlsTableStyle legacyStyle in workbook.TableStyles) {
                DocumentFormat.OpenXml.Spreadsheet.TableStyle? tableStyle = ProjectTableStyle(legacyStyle);
                if (tableStyle != null) {
                    tableStyles.Append(tableStyle);
                }
            }

            tableStyles.Count = checked((uint)tableStyles.Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyle>().Count());
            stylesheet.Save();
        }

        private static DocumentFormat.OpenXml.Spreadsheet.TableStyle? ProjectTableStyle(LegacyXlsTableStyle legacyStyle) {
            if (string.IsNullOrWhiteSpace(legacyStyle.Name)) {
                return null;
            }

            var tableStyle = new DocumentFormat.OpenXml.Spreadsheet.TableStyle {
                Name = legacyStyle.Name,
                Table = legacyStyle.AppliesToTables,
                Pivot = legacyStyle.AppliesToPivotTables
            };

            foreach (LegacyXlsTableStyleElement legacyElement in legacyStyle.Elements) {
                if (!TryMapTableStyleElementType(legacyElement.ElementType, out TableStyleValues elementType)) {
                    return null;
                }

                var element = new TableStyleElement {
                    Type = elementType,
                    FormatId = legacyElement.DifferentialFormatIndex
                };
                if (legacyElement.StripeSize != 0) {
                    element.Size = legacyElement.StripeSize;
                }

                tableStyle.Append(element);
            }

            tableStyle.Count = checked((uint)tableStyle.Elements<TableStyleElement>().Count());
            return tableStyle;
        }

        private static bool TryMapTableStyleElementType(uint value, out TableStyleValues type) {
            switch (value) {
                case 0x00000000:
                    type = TableStyleValues.WholeTable;
                    return true;
                case 0x00000001:
                    type = TableStyleValues.HeaderRow;
                    return true;
                case 0x00000002:
                    type = TableStyleValues.TotalRow;
                    return true;
                case 0x00000003:
                    type = TableStyleValues.FirstColumn;
                    return true;
                case 0x00000004:
                    type = TableStyleValues.LastColumn;
                    return true;
                case 0x00000005:
                    type = TableStyleValues.FirstRowStripe;
                    return true;
                case 0x00000006:
                    type = TableStyleValues.SecondRowStripe;
                    return true;
                case 0x00000007:
                    type = TableStyleValues.FirstColumnStripe;
                    return true;
                case 0x00000008:
                    type = TableStyleValues.SecondColumnStripe;
                    return true;
                case 0x00000009:
                    type = TableStyleValues.FirstHeaderCell;
                    return true;
                case 0x0000000A:
                    type = TableStyleValues.LastHeaderCell;
                    return true;
                case 0x0000000B:
                    type = TableStyleValues.FirstTotalCell;
                    return true;
                case 0x0000000C:
                    type = TableStyleValues.LastTotalCell;
                    return true;
                case 0x0000000D:
                    type = TableStyleValues.FirstSubtotalColumn;
                    return true;
                case 0x0000000E:
                    type = TableStyleValues.SecondSubtotalColumn;
                    return true;
                case 0x0000000F:
                    type = TableStyleValues.ThirdSubtotalColumn;
                    return true;
                case 0x00000010:
                    type = TableStyleValues.FirstSubtotalRow;
                    return true;
                case 0x00000011:
                    type = TableStyleValues.SecondSubtotalRow;
                    return true;
                case 0x00000012:
                    type = TableStyleValues.ThirdSubtotalRow;
                    return true;
                case 0x00000013:
                    type = TableStyleValues.BlankRow;
                    return true;
                case 0x00000014:
                    type = TableStyleValues.FirstColumnSubheading;
                    return true;
                case 0x00000015:
                    type = TableStyleValues.SecondColumnSubheading;
                    return true;
                case 0x00000016:
                    type = TableStyleValues.ThirdColumnSubheading;
                    return true;
                case 0x00000017:
                    type = TableStyleValues.FirstRowSubheading;
                    return true;
                case 0x00000018:
                    type = TableStyleValues.SecondRowSubheading;
                    return true;
                case 0x00000019:
                    type = TableStyleValues.ThirdRowSubheading;
                    return true;
                case 0x0000001A:
                    type = TableStyleValues.PageFieldLabels;
                    return true;
                case 0x0000001B:
                    type = TableStyleValues.PageFieldValues;
                    return true;
                default:
                    type = default;
                    return false;
            }
        }

        private static void ProjectWriteReservation(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (workbook.WriteReservation == null) {
                return;
            }

            document.SetWriteReservation(new ExcelWorkbookWriteReservationOptions {
                ReadOnlyRecommended = workbook.WriteReservation.ReadOnlyRecommended,
                UserName = workbook.WriteReservation.UserName
                    ?? (string.IsNullOrWhiteSpace(workbook.WriteReservation.LegacyPasswordHash) ? workbook.LastWriteUserName : null),
                LegacyPasswordHash = workbook.WriteReservation.LegacyPasswordHash
            });
        }

        private static void ProjectWorkbookRevisionProtection(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (!workbook.RevisionTrackingLocked.HasValue && !workbook.RevisionTrackingPasswordHash.HasValue) {
                return;
            }

            WorkbookProtection protection = GetOrCreateWorkbookProtection(document);
            if (workbook.RevisionTrackingLocked.HasValue) {
                protection.LockRevision = workbook.RevisionTrackingLocked.Value;
            }

            if (workbook.RevisionTrackingPasswordHash.HasValue) {
                protection.RevisionsPassword = workbook.RevisionTrackingPasswordHash.Value.ToString("X4", CultureInfo.InvariantCulture);
            }
        }

        private static WorkbookProtection GetOrCreateWorkbookProtection(ExcelDocument document) {
            Workbook workbook = document.WorkbookRoot;
            WorkbookProtection? protection = workbook.GetFirstChild<WorkbookProtection>();
            if (protection != null) {
                return protection;
            }

            protection = new WorkbookProtection();
            OpenXmlElement? before = workbook.GetFirstChild<BookViews>();
            before ??= workbook.GetFirstChild<Sheets>();
            if (before != null) {
                workbook.InsertBefore(protection, before);
            } else if (workbook.GetFirstChild<WorkbookProperties>() is WorkbookProperties workbookProperties) {
                workbook.InsertAfter(protection, workbookProperties);
            } else if (workbook.GetFirstChild<FileSharing>() is FileSharing fileSharing) {
                workbook.InsertAfter(protection, fileSharing);
            } else if (workbook.GetFirstChild<FileVersion>() is FileVersion fileVersion) {
                workbook.InsertAfter(protection, fileVersion);
            } else {
                workbook.InsertAt(protection, 0);
            }

            return protection;
        }

        private static void ProjectWorkbookOptions(LegacyXlsWorkbook workbook, ExcelDocument document) {
            WorkbookProperties? properties = null;
            if (workbook.SaveBackup.HasValue) {
                properties ??= GetOrCreateWorkbookProperties(document);
                properties.BackupFile = workbook.SaveBackup.Value;
            }

            if (workbook.DoNotSaveExternalLinkValues.HasValue) {
                properties ??= GetOrCreateWorkbookProperties(document);
                properties.SaveExternalLinkValues = !workbook.DoNotSaveExternalLinkValues.Value;
            }

            if (workbook.HiddenObjectsMode.HasValue && TryMapObjectDisplay(workbook.HiddenObjectsMode.Value, out ObjectDisplayValues objectDisplay)) {
                properties ??= GetOrCreateWorkbookProperties(document);
                properties.ShowObjects = objectDisplay;
            }

            if (workbook.HideBordersForInactiveTables.HasValue) {
                properties ??= GetOrCreateWorkbookProperties(document);
                properties.ShowBorderUnselectedTables = !workbook.HideBordersForInactiveTables.Value;
            }

            if (workbook.HasRefreshAllMarker) {
                properties ??= GetOrCreateWorkbookProperties(document);
                properties.RefreshAllConnections = true;
            }
        }

        private static void ProjectCodeNames(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (!string.IsNullOrWhiteSpace(workbook.CodeName)) {
                WorkbookProperties properties = GetOrCreateWorkbookProperties(document);
                properties.CodeName = workbook.CodeName;
            }

            for (int i = 0; i < workbook.Worksheets.Count && i < document.Sheets.Count; i++) {
                string? codeName = workbook.Worksheets[i].CodeName;
                if (string.IsNullOrWhiteSpace(codeName)) {
                    continue;
                }

                Worksheet worksheet = document.Sheets[i].WorksheetPart.Worksheet!;
                SheetProperties properties = worksheet.GetFirstChild<SheetProperties>() ?? new SheetProperties();
                if (properties.Parent == null) {
                    SheetDimension? dimension = worksheet.GetFirstChild<SheetDimension>();
                    if (dimension != null) {
                        worksheet.InsertBefore(properties, dimension);
                    } else {
                        worksheet.InsertAt(properties, 0);
                    }
                }

                properties.CodeName = codeName;
            }
        }

        private static WorkbookProperties GetOrCreateWorkbookProperties(ExcelDocument document) {
            WorkbookProperties properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>() ?? new WorkbookProperties();
            if (properties.Parent == null) {
                OpenXmlWorkbookElementOrder.InsertInOrder(document.WorkbookRoot, properties);
            }

            return properties;
        }

        private static bool TryMapObjectDisplay(ushort hiddenObjectsMode, out ObjectDisplayValues objectDisplay) {
            switch (hiddenObjectsMode) {
                case 0:
                    objectDisplay = ObjectDisplayValues.All;
                    return true;
                case 1:
                    objectDisplay = ObjectDisplayValues.Placeholders;
                    return true;
                case 2:
                    objectDisplay = ObjectDisplayValues.None;
                    return true;
                default:
                    objectDisplay = ObjectDisplayValues.All;
                    return false;
            }
        }

        private static void ProjectSheetTabIds(LegacyXlsWorkbook workbook, ExcelDocument document) {
            IReadOnlyList<ushort>? tabIds = workbook.SheetTabIds?.TabIds;
            if (tabIds == null || tabIds.Count != workbook.Worksheets.Count || tabIds.Any(id => id == 0) || tabIds.Distinct().Count() != tabIds.Count) {
                return;
            }

            Sheet[] sheets = document.WorkbookRoot.Sheets?.Elements<Sheet>().ToArray() ?? Array.Empty<Sheet>();
            if (sheets.Length != tabIds.Count) {
                return;
            }

            for (int i = 0; i < sheets.Length; i++) {
                sheets[i].SheetId = tabIds[i];
            }
        }

        private static void ProjectCalculationSettings(LegacyXlsWorkbook workbook, ExcelDocument document) {
            LegacyXlsCalculationSettings settings = workbook.CalculationSettings;
            if (settings.Records.Count == 0) {
                return;
            }

            CalculationProperties properties = document.WorkbookRoot.GetFirstChild<CalculationProperties>() ?? new CalculationProperties();
            if (properties.Parent == null) {
                OpenXmlWorkbookElementOrder.InsertInOrder(document.WorkbookRoot, properties);
            }

            if (settings.Mode.HasValue) {
                properties.CalculationMode = settings.Mode.Value switch {
                    LegacyXlsCalculationMode.Manual => CalculateModeValues.Manual,
                    LegacyXlsCalculationMode.AutomaticExceptTables => CalculateModeValues.AutoNoTable,
                    _ => CalculateModeValues.Auto
                };
            }

            if (settings.IterationCount.HasValue && settings.IterationCount.Value >= 0) {
                properties.IterateCount = checked((uint)settings.IterationCount.Value);
            }

            if (settings.FullPrecision.HasValue) {
                properties.FullPrecision = settings.FullPrecision.Value;
            }

            if (settings.A1ReferenceMode.HasValue) {
                properties.ReferenceMode = settings.A1ReferenceMode.Value
                    ? ReferenceModeValues.A1
                    : ReferenceModeValues.R1C1;
            }

            if (settings.Delta.HasValue && settings.Delta.Value >= 0d && !double.IsNaN(settings.Delta.Value) && !double.IsInfinity(settings.Delta.Value)) {
                properties.IterateDelta = settings.Delta.Value;
            }

            if (settings.IterationEnabled.HasValue) {
                properties.Iterate = settings.IterationEnabled.Value;
            }

            if (settings.RecalculateBeforeSave.HasValue) {
                properties.CalculationOnSave = settings.RecalculateBeforeSave.Value;
            }
        }

        private static void ProjectWorkbookWindow(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (workbook.Windows.Count == 0) {
                return;
            }

            LegacyXlsWorkbookWindow window = workbook.Windows[0];
            if (window.ActiveSheetIndex < workbook.Worksheets.Count) {
                LegacyXlsWorksheet activeSheet = workbook.Worksheets[window.ActiveSheetIndex];
                if (activeSheet.Visibility == 0) {
                    document.SetActiveWorksheet(window.ActiveSheetIndex);
                }
            }

            BookViews workbookViews = GetOrCreateWorkbookViews(document);
            workbookViews.RemoveAllChildren<WorkbookView>();
            foreach (LegacyXlsWorkbookWindow legacyWindow in workbook.Windows) {
                WorkbookView workbookView = workbookViews.AppendChild(new WorkbookView());
                ProjectWorkbookWindowView(workbook, legacyWindow, workbookView);
            }
        }

        private static void ProjectWorkbookWindowView(LegacyXlsWorkbook workbook, LegacyXlsWorkbookWindow window, WorkbookView workbookView) {
            workbookView.XWindow = window.HorizontalPositionTwips;
            workbookView.YWindow = window.VerticalPositionTwips;
            workbookView.WindowWidth = checked((uint)Math.Max(0, (int)window.WidthTwips));
            workbookView.WindowHeight = checked((uint)Math.Max(0, (int)window.HeightTwips));
            workbookView.Visibility = ToWorkbookViewVisibility(window);
            workbookView.Minimized = window.Minimized;
            workbookView.ActiveTab = window.ActiveSheetIndex < workbook.Worksheets.Count
                ? window.ActiveSheetIndex
                : 0U;
            workbookView.FirstSheet = window.FirstVisibleSheetTabIndex < workbook.Worksheets.Count
                ? (uint)window.FirstVisibleSheetTabIndex
                : window.ActiveSheetIndex < workbook.Worksheets.Count ? window.ActiveSheetIndex : 0U;
            workbookView.ShowHorizontalScroll = window.HorizontalScrollBarVisible;
            workbookView.ShowVerticalScroll = window.VerticalScrollBarVisible;
            workbookView.ShowSheetTabs = window.SheetTabsVisible;
            workbookView.TabRatio = window.SheetTabRatio;
        }

        private static BookViews GetOrCreateWorkbookViews(ExcelDocument document) {
            BookViews? workbookViews = document.WorkbookRoot.GetFirstChild<BookViews>();
            if (workbookViews == null) {
                workbookViews = new BookViews();
                Sheets? sheets = document.WorkbookRoot.GetFirstChild<Sheets>();
                if (sheets != null) {
                    document.WorkbookRoot.InsertBefore(workbookViews, sheets);
                } else {
                    document.WorkbookRoot.Append(workbookViews);
                }
            }

            return workbookViews;
        }

        private static WorkbookView GetOrCreatePrimaryWorkbookView(ExcelDocument document) {
            BookViews workbookViews = GetOrCreateWorkbookViews(document);
            WorkbookView? workbookView = workbookViews.Elements<WorkbookView>().FirstOrDefault();
            if (workbookView == null) {
                workbookView = new WorkbookView();
                workbookViews.Append(workbookView);
            }

            return workbookView;
        }

        private static SheetView GetOrCreatePrimarySheetView(ExcelSheet sheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetViews sheetViews = GetOrCreateSheetViews(worksheet);

            SheetView? sheetView = sheetViews.Elements<SheetView>().FirstOrDefault();
            if (sheetView == null) {
                sheetView = new SheetView { WorkbookViewId = 0U };
                sheetViews.Append(sheetView);
            } else if (sheetView.WorkbookViewId == null) {
                sheetView.WorkbookViewId = 0U;
            }

            return sheetView;
        }

        private static SheetViews GetOrCreateSheetViews(Worksheet worksheet) {
            SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                sheetViews = new SheetViews();
                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData != null) {
                    worksheet.InsertBefore(sheetViews, sheetData);
                } else {
                    worksheet.Append(sheetViews);
                }
            }

            return sheetViews;
        }

        private static void ProjectWorksheetWindowViews(ExcelSheet sheet, LegacyXlsWorksheet legacySheet) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetViews sheetViews = GetOrCreateSheetViews(worksheet);
            SheetView? existingPrimaryView = sheetViews.Elements<SheetView>().FirstOrDefault();
            Pane? primaryPane = existingPrimaryView?.GetFirstChild<Pane>()?.CloneNode(true) as Pane;

            sheetViews.RemoveAllChildren<SheetView>();
            for (int index = 0; index < legacySheet.WindowViews.Count; index++) {
                SheetView sheetView = CreateSheetView(legacySheet.WindowViews[index], checked((uint)index));
                if (index == 0 && primaryPane != null) {
                    sheetView.PrependChild(primaryPane);
                }

                sheetViews.Append(sheetView);
            }
        }

        private static SheetView CreateSheetView(LegacyXlsWorksheetWindowView view, uint workbookViewId) {
            var sheetView = new SheetView {
                WorkbookViewId = workbookViewId,
                ShowFormulas = view.ShowFormulas,
                ShowGridLines = view.ShowGridLines,
                ShowRowColHeaders = view.ShowRowColumnHeadings,
                ShowZeros = view.ShowZeroValues,
                RightToLeft = view.RightToLeft,
                DefaultGridColor = view.DefaultGridColor,
                ShowOutlineSymbols = view.ShowOutlineSymbols,
                TabSelected = view.TabSelected,
                View = view.PageBreakPreview ? SheetViewValues.PageBreakPreview : SheetViewValues.Normal
            };

            if (!view.DefaultGridColor && view.GridLineColorIndex.HasValue) {
                sheetView.ColorId = view.GridLineColorIndex.Value;
            }

            if (view.FirstVisibleRow.HasValue && view.FirstVisibleColumn.HasValue) {
                sheetView.TopLeftCell = A1.CellReference(view.FirstVisibleRow.Value + 1, view.FirstVisibleColumn.Value + 1);
            }

            if (view.ZoomScale.HasValue) {
                sheetView.ZoomScale = view.ZoomScale.Value;
            } else if (!view.PageBreakPreview && view.ZoomScaleNormal.HasValue) {
                sheetView.ZoomScale = view.ZoomScaleNormal.Value;
            }

            if (view.ZoomScaleNormal.HasValue) {
                sheetView.ZoomScaleNormal = view.ZoomScaleNormal.Value;
            }

            return sheetView;
        }

        private static VisibilityValues ToWorkbookViewVisibility(LegacyXlsWorkbookWindow window) {
            if (window.VeryHidden) {
                return VisibilityValues.VeryHidden;
            }

            return window.Hidden
                ? VisibilityValues.Hidden
                : VisibilityValues.Visible;
        }

        private static void ProjectDocumentProperties(LegacyXlsWorkbook workbook, ExcelDocument document) {
            LegacyXlsDocumentProperties? properties = workbook.DocumentProperties;
            if (properties == null || !properties.HasAnyProperties) {
                return;
            }

            if (properties.HasBuiltInProperties) {
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
            }

            if (!string.IsNullOrEmpty(properties.Company)) {
                document.ApplicationProperties.Company = properties.Company!;
            }

            if (!string.IsNullOrEmpty(properties.Manager)) {
                document.ApplicationProperties.Manager = properties.Manager!;
            }

            foreach (KeyValuePair<string, LegacyXlsDocumentPropertyValue> property in properties.CustomProperties) {
                document.SetCustomDocumentProperty(property.Key, ToExcelCustomProperty(property.Value));
            }
        }

        private static ExcelCustomProperty ToExcelCustomProperty(LegacyXlsDocumentPropertyValue property) {
            return property.Kind switch {
                LegacyXlsDocumentPropertyValueKind.Boolean => new ExcelCustomProperty(Convert.ToBoolean(property.Value, CultureInfo.InvariantCulture)),
                LegacyXlsDocumentPropertyValueKind.DateTime => new ExcelCustomProperty(Convert.ToDateTime(property.Value, CultureInfo.InvariantCulture)),
                LegacyXlsDocumentPropertyValueKind.Integer => new ExcelCustomProperty(property.Value, ExcelCustomPropertyType.NumberInteger),
                LegacyXlsDocumentPropertyValueKind.Number => new ExcelCustomProperty(Convert.ToDouble(property.Value, CultureInfo.InvariantCulture)),
                LegacyXlsDocumentPropertyValueKind.Binary => new ExcelCustomProperty((byte[])property.Value!),
                _ => new ExcelCustomProperty(Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty)
            };
        }

        private static void ProjectWorksheet(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            sheet.Batch(currentSheet => {
                foreach (LegacyXlsCell cell in legacySheet.Cells) {
                    LegacyXlsCellFormat? format = workbook.GetEffectiveCellFormat(cell.StyleIndex);
                    if (cell.Kind == LegacyXlsCellValueKind.Blank) {
                        ApplyCellFormat(currentSheet, workbook, cell, format);
                        continue;
                    }

                    object? value = GetProjectedCellValue(workbook, cell, format);
                    if (cell.Kind == LegacyXlsCellValueKind.Text
                        && value is string text
                        && TryCreateCellRichTextRuns(workbook, text, cell.TextFormattingRuns, out IReadOnlyList<ExcelRichTextRun> richTextRuns)) {
                        currentSheet.SetRichText(cell.Row, cell.Column, richTextRuns);
                    } else if (cell.Kind == LegacyXlsCellValueKind.Error && value is string errorText) {
                        currentSheet.SetLegacyErrorCellValue(cell.Row, cell.Column, errorText);
                    } else {
                        currentSheet.CellValue(cell.Row, cell.Column, value);
                    }

                    if (cell.IsFormula &&
                        !string.IsNullOrWhiteSpace(cell.FormulaText) &&
                        ShouldProjectFormula(workbook, cell.FormulaText!)) {
                        currentSheet.CellFormula(cell.Row, cell.Column, cell.FormulaText!);
                    }

                    ApplyCellFormat(currentSheet, workbook, cell, format);
                }

                foreach (LegacyXlsArrayFormulaRecord arrayFormula in legacySheet.ArrayFormulaRecords) {
                    LegacyXlsCell? formulaCell = legacySheet.Cells.FirstOrDefault(cell =>
                        cell.Row == arrayFormula.FirstRow
                        && cell.Column == arrayFormula.FirstColumn
                        && cell.IsFormula
                        && !string.IsNullOrWhiteSpace(cell.FormulaText)
                        && ShouldProjectFormula(workbook, cell.FormulaText!));
                    if (formulaCell != null) {
                        currentSheet.SetLegacyArrayFormula(arrayFormula.Range, formulaCell.FormulaText!);
                    }
                }

                if (legacySheet.DefaultColumnWidth.HasValue) {
                    currentSheet.SetDefaultColumnWidth(legacySheet.DefaultColumnWidth.Value, save: false);
                }

                foreach (LegacyXlsColumnLayout column in legacySheet.Columns) {
                    LegacyXlsCellFormat? columnFormat = workbook.GetEffectiveCellFormat(column.StyleIndex);
                    uint? projectedColumnStyleIndex = columnFormat != null
                        ? currentSheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, columnFormat)
                        : null;

                    for (int columnIndex = column.StartColumn; columnIndex <= column.EndColumn; columnIndex++) {
                        if (column.Width > 0) {
                            currentSheet.SetColumnWidth(columnIndex, column.Width);
                        }

                        if (column.Hidden) {
                            currentSheet.SetColumnHidden(columnIndex, true);
                        }

                        if (column.OutlineLevel > 0 || column.Collapsed) {
                            currentSheet.SetColumnOutline(columnIndex, column.OutlineLevel, column.Collapsed, save: false);
                        }

                        if (projectedColumnStyleIndex.HasValue && projectedColumnStyleIndex.Value != 0U) {
                            currentSheet.SetColumnStyleIndex(columnIndex, projectedColumnStyleIndex.Value, save: false);
                        }
                    }
                }

                if (legacySheet.DefaultRowHeight.HasValue) {
                    currentSheet.SetDefaultRowHeight(legacySheet.DefaultRowHeight.Value, legacySheet.DefaultRowsHidden, save: false);
                }

                foreach (LegacyXlsRowLayout row in legacySheet.Rows) {
                    if (row.CustomHeight && row.Height > 0) {
                        currentSheet.SetRowHeight(row.Row, row.Height);
                    }

                    if (row.Hidden) {
                        currentSheet.SetRowHidden(row.Row, true);
                    }

                    if (row.OutlineLevel > 0 || row.Collapsed) {
                        currentSheet.SetRowOutline(row.Row, row.OutlineLevel, row.Collapsed, save: false);
                    }

                    if (row.StyleIndex.HasValue) {
                        LegacyXlsCellFormat? rowFormat = workbook.GetEffectiveCellFormat(row.StyleIndex.Value);
                        if (rowFormat != null) {
                            uint projectedRowStyleIndex = currentSheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, rowFormat);
                            if (projectedRowStyleIndex != 0U) {
                                currentSheet.SetRowStyleIndex(row.Row, projectedRowStyleIndex, save: false);
                            }
                        }
                    }
                }

                foreach (LegacyXlsMergedRange mergedRange in legacySheet.MergedRanges) {
                    currentSheet.MergeRange(ToA1Range(mergedRange));
                }

                if (legacySheet.FreezePane != null) {
                    currentSheet.Freeze(legacySheet.FreezePane.TopRows, legacySheet.FreezePane.LeftColumns);
                    if (legacySheet.FrozenWithoutSplit == true) {
                        ApplyFrozenWithoutSplit(currentSheet);
                    }
                } else if (legacySheet.SplitPane != null) {
                    ApplySplitPane(currentSheet, legacySheet.SplitPane);
                }

                if (legacySheet.WindowViews.Count > 1) {
                    ProjectWorksheetWindowViews(currentSheet, legacySheet);
                } else {
                    SheetView? sheetView = null;

                    if (legacySheet.ZoomScale.HasValue) {
                        currentSheet.SetZoomScale(legacySheet.ZoomScale.Value, save: false);
                    }

                    if (legacySheet.ZoomScaleNormal.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.ZoomScaleNormal = legacySheet.ZoomScaleNormal.Value;
                    }

                    if (legacySheet.FirstVisibleRow.HasValue && legacySheet.FirstVisibleColumn.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.TopLeftCell = A1.CellReference(legacySheet.FirstVisibleRow.Value + 1, legacySheet.FirstVisibleColumn.Value + 1);
                    }

                    if (legacySheet.ShowGridLines.HasValue) {
                        currentSheet.SetGridlinesVisible(legacySheet.ShowGridLines.Value);
                    }

                    if (legacySheet.ShowFormulas.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.ShowFormulas = legacySheet.ShowFormulas.Value;
                    }

                    if (legacySheet.ShowRowColumnHeadings.HasValue) {
                        currentSheet.SetRowColumnHeadingsVisible(legacySheet.ShowRowColumnHeadings.Value);
                    }

                    if (legacySheet.ShowZeroValues.HasValue) {
                        currentSheet.SetZeroValuesVisible(legacySheet.ShowZeroValues.Value);
                    }

                    if (legacySheet.RightToLeft.HasValue) {
                        currentSheet.SetRightToLeft(legacySheet.RightToLeft.Value);
                    }

                    if (legacySheet.DefaultGridColor.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.DefaultGridColor = legacySheet.DefaultGridColor.Value;
                    }

                    if (legacySheet.GridLineColorIndex.HasValue && legacySheet.DefaultGridColor == false) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.ColorId = legacySheet.GridLineColorIndex.Value;
                    }

                    if (legacySheet.ShowOutlineSymbols.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.ShowOutlineSymbols = legacySheet.ShowOutlineSymbols.Value;
                    }

                    if (legacySheet.TabSelected.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.TabSelected = legacySheet.TabSelected.Value;
                    }

                    if (legacySheet.PageBreakPreview.HasValue) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.View = legacySheet.PageBreakPreview.Value
                            ? SheetViewValues.PageBreakPreview
                            : SheetViewValues.Normal;
                    }

                    if (legacySheet.PageLayoutView == true) {
                        sheetView ??= GetOrCreatePrimarySheetView(currentSheet);
                        sheetView.View = SheetViewValues.PageLayout;
                        if (legacySheet.PageLayoutZoomScale.HasValue) {
                            sheetView.ZoomScale = legacySheet.PageLayoutZoomScale.Value;
                        }
                    }
                }

                if (legacySheet.TabColorIndex.HasValue
                    && workbook.TryResolveColor(legacySheet.TabColorIndex.Value, out string? tabColorArgb)
                    && !string.IsNullOrWhiteSpace(tabColorArgb)) {
                    ApplyTabColor(currentSheet, tabColorArgb!);
                }

                foreach (LegacyXlsSelection selection in legacySheet.Selections) {
                    ProjectSelection(currentSheet, selection, legacySheet.FreezePane != null || legacySheet.SplitPane != null);
                }

                ProjectSortSettings(currentSheet, legacySheet.SortSettings);

                foreach (LegacyXlsHyperlink hyperlink in legacySheet.Hyperlinks) {
                    string reference = ToA1Range(hyperlink);
                    if (!string.IsNullOrWhiteSpace(hyperlink.DisplayText)) {
                        currentSheet.CellValue(hyperlink.StartRow, hyperlink.StartColumn, hyperlink.DisplayText!);
                    }

                    if (hyperlink.IsExternal) {
                        currentSheet.AddExternalHyperlinkReference(reference, hyperlink.Target, tooltip: hyperlink.Tooltip);
                    } else {
                        currentSheet.AddInternalHyperlinkReference(reference, hyperlink.Target, normalizeLocation: false, tooltip: hyperlink.Tooltip);
                    }
                }

                foreach (LegacyXlsDataValidation validation in legacySheet.DataValidations) {
                    LegacyXlsDataValidationProjector.Project(workbook, currentSheet, validation);
                }

                foreach (LegacyXlsConditionalFormatting conditionalFormatting in legacySheet.ConditionalFormattings) {
                    LegacyXlsConditionalFormattingProjector.Project(currentSheet, conditionalFormatting);
                }
            });

            ProjectTableDefinitions(workbook, legacySheet, sheet);

            foreach (LegacyXlsComment comment in legacySheet.Comments) {
                ExcelCommentAnchor? anchor = ToCommentAnchor(comment.Anchor);
                if (TryCreateCommentRichTextRuns(workbook, comment, out IReadOnlyList<ExcelRichTextRun> richTextRuns)) {
                    sheet.SetLegacyCommentRichText(comment.Row, comment.Column, richTextRuns, comment.Author, comment.Visible, anchor);
                } else {
                    sheet.SetLegacyComment(comment.Row, comment.Column, comment.Text, comment.Author, comment.Visible, anchor);
                }
            }

            if (legacySheet.Protection?.IsProtected == true) {
                LegacyXlsWorksheetProtectionPermissions? permissions = legacySheet.Protection.Permissions;
                sheet.Protect(new ExcelSheetProtectionOptions {
                    LegacyPasswordHash = legacySheet.Protection.LegacyPasswordHash,
                    ProtectObjects = permissions != null ? !permissions.AllowEditObjects : legacySheet.Protection.ProtectObjects,
                    ProtectScenarios = permissions != null ? !permissions.AllowEditScenarios : legacySheet.Protection.ProtectScenarios,
                    AllowSelectLockedCells = permissions?.AllowSelectLockedCells ?? true,
                    AllowSelectUnlockedCells = permissions?.AllowSelectUnlockedCells ?? true,
                    AllowFormatCells = permissions?.AllowFormatCells ?? false,
                    AllowFormatColumns = permissions?.AllowFormatColumns ?? false,
                    AllowFormatRows = permissions?.AllowFormatRows ?? false,
                    AllowInsertColumns = permissions?.AllowInsertColumns ?? false,
                    AllowInsertRows = permissions?.AllowInsertRows ?? false,
                    AllowInsertHyperlinks = permissions?.AllowInsertHyperlinks ?? false,
                    AllowDeleteColumns = permissions?.AllowDeleteColumns ?? false,
                    AllowDeleteRows = permissions?.AllowDeleteRows ?? false,
                    AllowSort = permissions?.AllowSort ?? false,
                    AllowAutoFilter = permissions?.AllowAutoFilter ?? false,
                    AllowPivotTables = permissions?.AllowPivotTables ?? false
            });
        }

            ProjectProtectedRanges(legacySheet, sheet);
            ProjectPageSetup(legacySheet, sheet);
            ProjectWorksheetOptions(legacySheet, sheet);
            ProjectGridSet(legacySheet, sheet);
            ProjectWorksheetCalculationProperties(legacySheet, sheet);
            ProjectWorksheetPhoneticProperties(legacySheet, sheet);
            ProjectIgnoredErrors(legacySheet, sheet);
            ProjectCellWatches(legacySheet, sheet);
            ProjectDataConsolidation(workbook, legacySheet, sheet);
            ProjectScenarios(legacySheet, sheet);
            ProjectPageBreaks(legacySheet, sheet);
            ProjectWorksheetImages(workbook, legacySheet, sheet);
            ProjectHeaderFooterImages(workbook, legacySheet, sheet);

            if (legacySheet.Visibility == 2) {
                sheet.SetVeryHidden(true);
            } else if (legacySheet.Visibility != 0) {
                sheet.SetHidden(true);
            }
        }

        private static bool ShouldProjectFormula(LegacyXlsWorkbook workbook, string formulaText) =>
            workbook.PreserveExternalWorkbookLinks || !ReferencesExternalWorkbook(workbook, formulaText);

        private static bool ReferencesExternalWorkbook(LegacyXlsWorkbook workbook, string formulaText) {
            foreach (LegacyXlsExternalReference reference in workbook.ExternalReferences) {
                if (reference.Kind != LegacyXlsExternalReferenceKind.ExternalWorkbook ||
                    string.IsNullOrWhiteSpace(reference.Target)) {
                    continue;
                }

                string normalizedTarget = reference.Target!.Replace('\\', '/');
                int separator = normalizedTarget.LastIndexOf('/');
                string fileName = separator >= 0 ? normalizedTarget.Substring(separator + 1) : normalizedTarget;
                if (fileName.Length == 0) continue;
                string escapedFileName = fileName.Replace("'", "''");
                if (ContainsFormulaTokenOutsideStringLiteral(formulaText, "[" + escapedFileName + "]") ||
                    ContainsFormulaTokenOutsideStringLiteral(formulaText, "'" + escapedFileName + "'!") ||
                    ContainsFormulaTokenOutsideStringLiteral(formulaText, escapedFileName + "!")) {
                    return true;
                }
            }

            return false;
        }

        private static bool ContainsFormulaTokenOutsideStringLiteral(string formulaText, string token) {
            bool inStringLiteral = false;
            for (int index = 0; index < formulaText.Length;) {
                if (formulaText[index] == '"') {
                    if (inStringLiteral && index + 1 < formulaText.Length && formulaText[index + 1] == '"') {
                        index += 2;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    index++;
                    continue;
                }

                if (!inStringLiteral && index <= formulaText.Length - token.Length &&
                    string.Compare(formulaText, index, token, 0, token.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                    return true;
                }

                index++;
            }

            return false;
        }

        private static void ProjectTableDefinitions(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.TableDefinitions.Count == 0) {
                return;
            }

            OfficeIMO.Excel.TableStyle style = ResolveDefaultTableStyle(workbook);
            foreach (LegacyXlsTableDefinition tableDefinition in legacySheet.TableDefinitions) {
                sheet.AddTable(
                    tableDefinition.Range,
                    tableDefinition.HasHeaderRow,
                    tableDefinition.Name,
                    style,
                    tableDefinition.HasAutoFilter);
                ApplyTableDefinitionTotals(sheet, tableDefinition);
                ApplyTableDefinitionMetadata(sheet, tableDefinition);
                ApplyTableBlockLevelFormatting(workbook, sheet, tableDefinition);
            }
        }

        private static void ApplyTableDefinitionTotals(ExcelSheet sheet, LegacyXlsTableDefinition tableDefinition) {
            if (!tableDefinition.HasTotalsRow) {
                return;
            }

            Table? table = sheet.WorksheetPart.TableDefinitionParts
                .FirstOrDefault(part => string.Equals(part.Table?.Name?.Value, tableDefinition.Name, StringComparison.OrdinalIgnoreCase))
                ?.Table;
            if (table == null) {
                return;
            }

            table.TotalsRowShown = true;
            table.TotalsRowCount = tableDefinition.TotalRowCount;
            table.Save();
        }

        private static void ApplyTableDefinitionMetadata(ExcelSheet sheet, LegacyXlsTableDefinition tableDefinition) {
            if (string.IsNullOrWhiteSpace(tableDefinition.StyleName)
                && string.IsNullOrWhiteSpace(tableDefinition.DisplayName)
                && string.IsNullOrWhiteSpace(tableDefinition.Comment)
                && !tableDefinition.ShowFirstColumn.HasValue
                && !tableDefinition.ShowLastColumn.HasValue
                && !tableDefinition.ShowRowStripes.HasValue
                && !tableDefinition.ShowColumnStripes.HasValue) {
                return;
            }

            Table? table = sheet.WorksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .FirstOrDefault(candidate => string.Equals(candidate?.Name?.Value, tableDefinition.Name, StringComparison.OrdinalIgnoreCase));
            if (table == null) {
                return;
            }

            if (!string.IsNullOrWhiteSpace(tableDefinition.DisplayName)) {
                table.Name = tableDefinition.DisplayName;
                table.DisplayName = tableDefinition.DisplayName;
            }

            if (!string.IsNullOrWhiteSpace(tableDefinition.Comment)) {
                table.Comment = tableDefinition.Comment;
            }

            if (!string.IsNullOrWhiteSpace(tableDefinition.StyleName)
                || tableDefinition.ShowFirstColumn.HasValue
                || tableDefinition.ShowLastColumn.HasValue
                || tableDefinition.ShowRowStripes.HasValue
                || tableDefinition.ShowColumnStripes.HasValue) {
                TableStyleInfo styleInfo = table.TableStyleInfo ?? table.AppendChild(new TableStyleInfo());
                if (!string.IsNullOrWhiteSpace(tableDefinition.StyleName)) {
                    styleInfo.Name = tableDefinition.StyleName;
                }

                if (tableDefinition.ShowFirstColumn.HasValue) {
                    styleInfo.ShowFirstColumn = tableDefinition.ShowFirstColumn.Value;
                }

                if (tableDefinition.ShowLastColumn.HasValue) {
                    styleInfo.ShowLastColumn = tableDefinition.ShowLastColumn.Value;
                }

                if (tableDefinition.ShowRowStripes.HasValue) {
                    styleInfo.ShowRowStripes = tableDefinition.ShowRowStripes.Value;
                }

                if (tableDefinition.ShowColumnStripes.HasValue) {
                    styleInfo.ShowColumnStripes = tableDefinition.ShowColumnStripes.Value;
                }
            }

            table.Save();
        }

        private static OfficeIMO.Excel.TableStyle ResolveDefaultTableStyle(LegacyXlsWorkbook workbook) {
            string? defaultTableStyleName = workbook.TableStyleCollections.LastOrDefault()?.DefaultTableStyleName;
            if (!string.IsNullOrWhiteSpace(defaultTableStyleName)
                && Enum.TryParse(defaultTableStyleName, ignoreCase: true, out OfficeIMO.Excel.TableStyle tableStyle)) {
                return tableStyle;
            }

            return OfficeIMO.Excel.TableStyle.TableStyleMedium2;
        }

        private static void ProjectProtectedRanges(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.ProtectedRanges.Count == 0) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            ProtectedRanges protectedRanges = worksheet.GetFirstChild<ProtectedRanges>() ?? worksheet.AppendChild(new ProtectedRanges());
            foreach (LegacyXlsProtectedRange legacyRange in legacySheet.ProtectedRanges) {
                var protectedRange = new ProtectedRange {
                    Name = legacyRange.Name,
                    SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = string.Join(" ", legacyRange.References)
                    }
                };
                if (!string.IsNullOrWhiteSpace(legacyRange.LegacyPasswordHash)) {
                    protectedRange.Password = legacyRange.LegacyPasswordHash;
                }

                protectedRanges.Append(protectedRange);
            }

            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectIgnoredErrors(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.IgnoredErrors.Count == 0) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            IgnoredErrors ignoredErrors = worksheet.GetFirstChild<IgnoredErrors>() ?? worksheet.AppendChild(new IgnoredErrors());
            foreach (LegacyXlsIgnoredError legacyError in legacySheet.IgnoredErrors) {
                var ignoredError = new IgnoredError {
                    SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = string.Join(" ", legacyError.References)
                    }
                };
                if (legacyError.EvaluationError) ignoredError.EvalError = true;
                if (legacyError.EmptyCellReference) ignoredError.EmptyCellReference = true;
                if (legacyError.NumberStoredAsText) ignoredError.NumberStoredAsText = true;
                if (legacyError.FormulaRange) ignoredError.FormulaRange = true;
                if (legacyError.Formula) ignoredError.Formula = true;
                if (legacyError.TwoDigitTextYear) ignoredError.TwoDigitTextYear = true;
                if (legacyError.UnlockedFormula) ignoredError.UnlockedFormula = true;
                if (legacyError.ListDataValidation) ignoredError.ListDataValidation = true;
                ignoredErrors.Append(ignoredError);
            }

            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectCellWatches(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.CellWatches.Count == 0) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            CellWatches cellWatches = worksheet.GetFirstChild<CellWatches>() ?? worksheet.AppendChild(new CellWatches());
            foreach (LegacyXlsCellWatch legacyCellWatch in legacySheet.CellWatches) {
                cellWatches.Append(new CellWatch {
                    CellReference = legacyCellWatch.CellReference
                });
            }

            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectScenarios(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.Scenarios.Count == 0) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            Scenarios scenarios = worksheet.GetFirstChild<Scenarios>() ?? worksheet.AppendChild(new Scenarios());
            LegacyXlsScenarioManager? manager = legacySheet.ScenarioManager;
            if (manager != null) {
                if (manager.CurrentScenarioIndex >= 0) scenarios.Current = (uint)manager.CurrentScenarioIndex;
                if (manager.ShownScenarioIndex >= 0) scenarios.Show = (uint)manager.ShownScenarioIndex;
                if (manager.ResultRanges.Count > 0) {
                    scenarios.SequenceOfReferences = new ListValue<StringValue> {
                        InnerText = string.Join(" ", manager.ResultRanges)
                    };
                }
            }

            foreach (LegacyXlsScenario legacyScenario in legacySheet.Scenarios) {
                var scenario = new Scenario {
                    Name = legacyScenario.Name,
                    Count = (uint)legacyScenario.InputCells.Count
                };
                if (legacyScenario.Locked) scenario.Locked = true;
                if (legacyScenario.Hidden) scenario.Hidden = true;
                if (!string.IsNullOrEmpty(legacyScenario.User)) scenario.User = legacyScenario.User;
                if (!string.IsNullOrEmpty(legacyScenario.Comment)) scenario.Comment = legacyScenario.Comment;

                foreach (LegacyXlsScenarioInputCell inputCell in legacyScenario.InputCells) {
                    var openXmlInputCell = new InputCells {
                        CellReference = inputCell.CellReference,
                        Val = inputCell.Value
                    };
                    if (inputCell.Deleted) openXmlInputCell.Deleted = true;
                    scenario.Append(openXmlInputCell);
                }

                scenarios.Append(scenario);
            }

            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectDataConsolidation(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            LegacyXlsDataConsolidationSettings? legacySettings = legacySheet.DataConsolidationSettings;
            if (legacySettings == null) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            DataConsolidate dataConsolidate = worksheet.GetFirstChild<DataConsolidate>() ?? worksheet.AppendChild(new DataConsolidate());
            dataConsolidate.Function = ToDataConsolidateFunction(legacySettings.Function);
            if (legacySettings.UsesLeftLabels) dataConsolidate.LeftLabels = true;
            if (legacySettings.UsesTopLabels) dataConsolidate.TopLabels = true;
            if (legacySettings.LinksToSourceData) dataConsolidate.Link = true;
            IReadOnlyList<LegacyXlsDataConsolidationReference> references = GetProjectedDataConsolidationReferences(workbook, legacySheet);
            if (references.Count > 0) {
                DataReferences dataReferences = dataConsolidate.GetFirstChild<DataReferences>() ?? dataConsolidate.AppendChild(new DataReferences());
                foreach (LegacyXlsDataConsolidationReference reference in references) {
                    var dataReference = new DataReference {
                        Reference = reference.CellRange,
                    };
                    if (reference.SourceKind == LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath) {
                        if (!workbook.PreserveExternalWorkbookLinks) {
                            continue;
                        }

                        string? relationshipId = AddExternalDataConsolidationRelationship(sheet, reference.Source);
                        if (relationshipId == null) {
                            continue;
                        }

                        dataReference.Id = relationshipId;
                    } else {
                        dataReference.Sheet = reference.Source;
                    }

                    dataReferences.Append(dataReference);
                }

                dataReferences.Count = (uint)dataReferences.Elements<DataReference>().Count();
            }

            IReadOnlyList<LegacyXlsDataConsolidationName> names = GetProjectedDataConsolidationNames(workbook, legacySheet);
            if (names.Count > 0) {
                DataReferences dataReferences = dataConsolidate.GetFirstChild<DataReferences>() ?? dataConsolidate.AppendChild(new DataReferences());
                foreach (LegacyXlsDataConsolidationName name in names) {
                    var reference = new DataReference {
                        Name = name.Name
                    };
                    if (name.SourceKind == LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath) {
                        if (!workbook.PreserveExternalWorkbookLinks) {
                            continue;
                        }

                        string? relationshipId = AddExternalDataConsolidationRelationship(sheet, name.Source);
                        if (relationshipId == null) {
                            continue;
                        }

                        reference.Id = relationshipId;
                    } else if (!string.IsNullOrWhiteSpace(name.Source)) {
                        reference.Sheet = name.Source;
                    }

                    dataReferences.Append(reference);
                }

                dataReferences.Count = (uint)dataReferences.Elements<DataReference>().Count();
            }

            sheet.EnsureWorksheetElementOrder();
        }

        private static IReadOnlyList<LegacyXlsDataConsolidationReference> GetProjectedDataConsolidationReferences(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet) {
            if (legacySheet.DataConsolidationSettings == null) {
                return Array.Empty<LegacyXlsDataConsolidationReference>();
            }

            int ownerCount = workbook.Worksheets.Count(item => item.DataConsolidationSettings != null);
            if (ownerCount != 1) {
                return Array.Empty<LegacyXlsDataConsolidationReference>();
            }

            return workbook.DataConsolidationReferences
                .Where(reference => reference.SourceKind == LegacyXlsDataConsolidationSourceKind.SelfReference
                    || reference.SourceKind == LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath)
                .ToArray();
        }

        private static IReadOnlyList<LegacyXlsDataConsolidationName> GetProjectedDataConsolidationNames(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet) {
            if (legacySheet.DataConsolidationSettings == null) {
                return Array.Empty<LegacyXlsDataConsolidationName>();
            }

            int ownerCount = workbook.Worksheets.Count(item => item.DataConsolidationSettings != null);
            if (ownerCount != 1) {
                return Array.Empty<LegacyXlsDataConsolidationName>();
            }

            return workbook.DataConsolidationNames
                .Where(name => name.SourceKind == LegacyXlsDataConsolidationSourceKind.SelfReference
                    || name.SourceKind == LegacyXlsDataConsolidationSourceKind.ExternalVirtualPath)
                .ToArray();
        }

        private static string? AddExternalDataConsolidationRelationship(ExcelSheet sheet, string source) {
            if (!TryCreateExternalTargetUri(source, out Uri? targetUri) || targetUri == null) {
                return null;
            }

            ExternalRelationship relationship = sheet.WorksheetPart.AddExternalRelationship(
                ExternalLinkPathRelationshipType,
                targetUri);
            return relationship.Id;
        }

        private static DataConsolidateFunctionValues ToDataConsolidateFunction(LegacyXlsDataConsolidationFunction function) {
            switch (function) {
                case LegacyXlsDataConsolidationFunction.Average:
                    return DataConsolidateFunctionValues.Average;
                case LegacyXlsDataConsolidationFunction.CountNumbers:
                    return DataConsolidateFunctionValues.CountNumbers;
                case LegacyXlsDataConsolidationFunction.Count:
                    return DataConsolidateFunctionValues.Count;
                case LegacyXlsDataConsolidationFunction.Maximum:
                    return DataConsolidateFunctionValues.Maximum;
                case LegacyXlsDataConsolidationFunction.Minimum:
                    return DataConsolidateFunctionValues.Minimum;
                case LegacyXlsDataConsolidationFunction.Product:
                    return DataConsolidateFunctionValues.Product;
                case LegacyXlsDataConsolidationFunction.StandardDeviation:
                    return DataConsolidateFunctionValues.StandardDeviation;
                case LegacyXlsDataConsolidationFunction.StandardDeviationP:
                    return DataConsolidateFunctionValues.StandardDeviationP;
                case LegacyXlsDataConsolidationFunction.Variance:
                    return DataConsolidateFunctionValues.Variance;
                case LegacyXlsDataConsolidationFunction.VarianceP:
                    return DataConsolidateFunctionValues.VarianceP;
                default:
                    return DataConsolidateFunctionValues.Sum;
            }
        }

        private static void ProjectSelection(ExcelSheet sheet, LegacyXlsSelection selection, bool hasFrozenPane) {
            string activeCell = A1.CellReference(selection.ActiveRow, selection.ActiveColumn);
            IReadOnlyList<string> selectedRanges = selection.SelectedRanges
                .Select(range => range.Reference)
                .ToArray();
            sheet.SetWorksheetSelection(activeCell, selectedRanges, hasFrozenPane ? ToPane(selection.Pane) : null, save: false);
        }

        private static void ProjectSortSettings(ExcelSheet sheet, LegacyXlsSortSettings? sortSettings) {
            if (sortSettings == null) {
                return;
            }

            var conditions = new List<SortCondition>(3);
            int firstRow = int.MaxValue;
            int firstColumn = int.MaxValue;
            int lastRow = 0;
            int lastColumn = 0;

            if (!TryAppendSortCondition(conditions, sortSettings.Key1, sortSettings.Key1Descending, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn)
                || !TryAppendSortCondition(conditions, sortSettings.Key2, sortSettings.Key2Descending, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn)
                || !TryAppendSortCondition(conditions, sortSettings.Key3, sortSettings.Key3Descending, ref firstRow, ref firstColumn, ref lastRow, ref lastColumn)
                || conditions.Count == 0) {
                return;
            }

            var sortState = new SortState(conditions) {
                Reference = BuildA1Range(firstRow, firstColumn, lastRow, lastColumn)
            };
            if (sortSettings.SortLeftToRight) {
                sortState.ColumnSort = true;
            }

            if (sortSettings.CaseSensitive) {
                sortState.CaseSensitive = true;
            }

            if (sortSettings.UsePhoneticInformation) {
                sortState.SortMethod = SortMethodValues.PinYin;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.RemoveAllChildren<SortState>();
            worksheet.AppendChild(sortState);
            sheet.EnsureWorksheetElementOrder();
        }

        private static bool TryAppendSortCondition(
            ICollection<SortCondition> conditions,
            string? reference,
            bool descending,
            ref int firstRow,
            ref int firstColumn,
            ref int lastRow,
            ref int lastColumn) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return true;
            }

            if (!TryNormalizeSortReference(reference!, out string normalizedReference, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            firstRow = Math.Min(firstRow, r1);
            firstColumn = Math.Min(firstColumn, c1);
            lastRow = Math.Max(lastRow, r2);
            lastColumn = Math.Max(lastColumn, c2);

            var condition = new SortCondition {
                Reference = normalizedReference,
                SortBy = SortByValues.Value
            };
            if (descending) {
                condition.Descending = true;
            }

            conditions.Add(condition);
            return true;
        }

        private static bool TryNormalizeSortReference(string reference, out string normalizedReference, out int r1, out int c1, out int r2, out int c2) {
            normalizedReference = string.Empty;
            r1 = c1 = r2 = c2 = 0;

            string normalized = reference.Trim().Replace("$", string.Empty);
            if (normalized.IndexOfAny(new[] { '!', ',' }) >= 0 || normalized.Any(char.IsWhiteSpace)) {
                return false;
            }

            if (A1.TryParseRange(normalized, out r1, out c1, out r2, out c2)) {
                normalizedReference = BuildA1Range(r1, c1, r2, c2);
                return true;
            }

            if (A1.TryParseCellReferenceFast(normalized, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                normalizedReference = A1.CellReference(r1, c1);
                return true;
            }

            return false;
        }

        private static string BuildA1Range(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
        }

        private static PaneValues? ToPane(byte pane) {
            return pane switch {
                0x00 => PaneValues.BottomRight,
                0x01 => PaneValues.TopRight,
                0x02 => PaneValues.BottomLeft,
                0x03 => PaneValues.TopLeft,
                _ => null
            };
        }

        private static void ApplySplitPane(ExcelSheet sheet, LegacyXlsSplitPane splitPane) {
            SheetView sheetView = GetOrCreatePrimarySheetView(sheet);
            sheetView.RemoveAllChildren<Pane>();
            var pane = new Pane {
                State = PaneStateValues.Split,
                HorizontalSplit = splitPane.HorizontalSplit,
                VerticalSplit = splitPane.VerticalSplit,
                TopLeftCell = A1.CellReference(splitPane.TopRow + 1, splitPane.LeftColumn + 1)
            };

            PaneValues? activePane = ToPane(splitPane.ActivePane);
            if (activePane.HasValue) {
                pane.ActivePane = activePane.Value;
            }

            sheetView.PrependChild(pane);
        }

        private static void ApplyFrozenWithoutSplit(ExcelSheet sheet) {
            SheetView sheetView = GetOrCreatePrimarySheetView(sheet);
            Pane? pane = sheetView.GetFirstChild<Pane>();
            if (pane != null) {
                pane.State = PaneStateValues.FrozenSplit;
            }
        }

        private static void ApplyTabColor(ExcelSheet sheet, string argb) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetProperties properties = GetOrCreateSheetProperties(worksheet);

            properties.TabColor = new TabColor {
                Rgb = argb
            };
        }

        private static void ProjectWorksheetOptions(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (!legacySheet.ApplyOutlineStyles.HasValue
                && !legacySheet.SummaryRowsBelow.HasValue
                && !legacySheet.SummaryColumnsRightWhenLeftToRight.HasValue) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetProperties properties = GetOrCreateSheetProperties(worksheet);
            OutlineProperties outlineProperties = GetOrCreateOutlineProperties(properties);
            if (legacySheet.ApplyOutlineStyles.HasValue) {
                outlineProperties.ApplyStyles = legacySheet.ApplyOutlineStyles.Value;
            }

            if (legacySheet.SummaryRowsBelow.HasValue) {
                outlineProperties.SummaryBelow = legacySheet.SummaryRowsBelow.Value;
            }

            if (legacySheet.SummaryColumnsRightWhenLeftToRight.HasValue) {
                outlineProperties.SummaryRight = legacySheet.SummaryColumnsRightWhenLeftToRight.Value;
            }
        }

        private static void ProjectGridSet(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (!legacySheet.GridSet.HasValue) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            PrintOptions printOptions = worksheet.GetFirstChild<PrintOptions>() ?? worksheet.AppendChild(new PrintOptions());
            printOptions.GridLinesSet = legacySheet.GridSet.Value;
            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectWorksheetCalculationProperties(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            if (legacySheet.FullCalculationOnLoad != true) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            SheetCalculationProperties properties = worksheet.GetFirstChild<SheetCalculationProperties>() ?? worksheet.AppendChild(new SheetCalculationProperties());
            properties.FullCalculationOnLoad = true;
            sheet.EnsureWorksheetElementOrder();
        }

        private static void ProjectWorksheetPhoneticProperties(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            LegacyXlsPhoneticSettings? settings = legacySheet.PhoneticSettings;
            if (settings == null) {
                return;
            }

            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.RemoveAllChildren<PhoneticProperties>();
            worksheet.AppendChild(new PhoneticProperties {
                FontId = settings.FontId,
                Type = ToPhoneticValues(settings.Type),
                Alignment = ToPhoneticAlignmentValues(settings.Alignment)
            });
            sheet.EnsureWorksheetElementOrder();
        }

        private static PhoneticValues ToPhoneticValues(LegacyXlsPhoneticType value) {
            return value switch {
                LegacyXlsPhoneticType.HalfWidthKatakana => PhoneticValues.HalfWidthKatakana,
                LegacyXlsPhoneticType.FullWidthKatakana => PhoneticValues.FullWidthKatakana,
                LegacyXlsPhoneticType.Hiragana => PhoneticValues.Hiragana,
                LegacyXlsPhoneticType.NoConversion => PhoneticValues.NoConversion,
                _ => PhoneticValues.FullWidthKatakana
            };
        }

        private static PhoneticAlignmentValues ToPhoneticAlignmentValues(LegacyXlsPhoneticAlignment value) {
            return value switch {
                LegacyXlsPhoneticAlignment.NoControl => PhoneticAlignmentValues.NoControl,
                LegacyXlsPhoneticAlignment.Left => PhoneticAlignmentValues.Left,
                LegacyXlsPhoneticAlignment.Center => PhoneticAlignmentValues.Center,
                LegacyXlsPhoneticAlignment.Distributed => PhoneticAlignmentValues.Distributed,
                _ => PhoneticAlignmentValues.Left
            };
        }

        private static SheetProperties GetOrCreateSheetProperties(Worksheet worksheet) {
            SheetProperties properties = worksheet.GetFirstChild<SheetProperties>() ?? new SheetProperties();
            if (properties.Parent != null) {
                return properties;
            }

            SheetDimension? dimension = worksheet.GetFirstChild<SheetDimension>();
            if (dimension != null) {
                worksheet.InsertBefore(properties, dimension);
            } else {
                worksheet.InsertAt(properties, 0);
            }

            return properties;
        }

        private static OutlineProperties GetOrCreateOutlineProperties(SheetProperties properties) {
            OutlineProperties? outlineProperties = properties.GetFirstChild<OutlineProperties>();
            if (outlineProperties != null) {
                return outlineProperties;
            }

            outlineProperties = new OutlineProperties();
            PageSetupProperties? pageSetupProperties = properties.GetFirstChild<PageSetupProperties>();
            if (pageSetupProperties != null) {
                properties.InsertBefore(outlineProperties, pageSetupProperties);
            } else {
                properties.AppendChild(outlineProperties);
            }

            return outlineProperties;
        }

        private static ExcelCommentAnchor? ToCommentAnchor(LegacyXlsDrawingAnchor? anchor) {
            return anchor == null
                ? null
                : new ExcelCommentAnchor(
                    anchor.StartColumn,
                    anchor.StartDx,
                    anchor.StartRow,
                    anchor.StartDy,
                    anchor.EndColumn,
                    anchor.EndDx,
                    anchor.EndRow,
                    anchor.EndDy);
        }

        private static void ProjectDefinedNames(LegacyXlsWorkbook workbook, ExcelDocument document) {
            IReadOnlyDictionary<int, int> worksheetIndexByProjectedSheetIndex = CreateWorksheetIndexByProjectedSheetIndex(workbook);
            foreach (LegacyXlsDefinedName definedName in workbook.DefinedNames) {
                if (!workbook.PreserveExternalWorkbookLinks && ReferencesExternalWorkbook(workbook, definedName.Reference)) {
                    continue;
                }
                ExcelSheet? scope = definedName.LocalSheetIndex.HasValue
                    && worksheetIndexByProjectedSheetIndex.TryGetValue(definedName.LocalSheetIndex.Value, out int worksheetIndex)
                    && worksheetIndex < document.Sheets.Count
                    ? document.Sheets[worksheetIndex]
                    : null;
                if (definedName.LocalSheetIndex.HasValue && scope == null) {
                    AppendRawDefinedName(document, definedName);
                    continue;
                }

                if (scope != null
                    && string.Equals(definedName.Name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)
                    && TryParsePrintTitles(definedName.Reference, scope, out int? firstRow, out int? lastRow, out int? firstColumn, out int? lastColumn)) {
                    document.SetPrintTitles(scope, firstRow, lastRow, firstColumn, lastColumn, save: false);
                    continue;
                }

                if (scope != null
                    && string.Equals(definedName.Name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                    && TryParseScopedRange(definedName.Reference, scope, out string autoFilterRange)) {
                    scope.AddAutoFilter(autoFilterRange);
                }

                if (!TrySetNamedRange(document, definedName, scope)) {
                    AppendRawDefinedName(document, definedName);
                }
            }
        }

        private static bool TrySetNamedRange(ExcelDocument document, LegacyXlsDefinedName definedName, ExcelSheet? scope) {
            try {
                document.SetNamedRange(
                    definedName.Name,
                    definedName.Reference,
                    scope,
                    save: false,
                    hidden: definedName.Hidden,
                    validationMode: NameValidationMode.Strict);
                return true;
            } catch (ArgumentException) {
                return false;
            }
        }

        private static void AppendRawDefinedName(ExcelDocument document, LegacyXlsDefinedName definedName) {
            DefinedNames definedNames = document.WorkbookRoot.DefinedNames ??= new DefinedNames();
            var openXmlName = new DefinedName {
                Name = definedName.Name,
                Text = definedName.Reference,
                Hidden = definedName.Hidden ? true : (bool?)null
            };
            if (definedName.LocalSheetIndex.HasValue) {
                openXmlName.LocalSheetId = checked((uint)definedName.LocalSheetIndex.Value);
            }

            definedNames.Append(openXmlName);
        }

        private static void ProjectAutoFilters(LegacyXlsWorkbook workbook, ExcelDocument document) {
            IReadOnlyDictionary<int, int> worksheetIndexByProjectedSheetIndex = CreateWorksheetIndexByProjectedSheetIndex(workbook);
            for (int i = 0; i < workbook.Worksheets.Count && i < document.Sheets.Count; i++) {
                LegacyXlsWorksheet legacySheet = workbook.Worksheets[i];
                if (legacySheet.AutoFilterCriteria.Count == 0) {
                    continue;
                }

                ExcelSheet sheet = document.Sheets[i];
                if (!TryGetAutoFilterRange(workbook, sheet, i, worksheetIndexByProjectedSheetIndex, out string autoFilterRange)) {
                    continue;
                }

                LegacyXlsAutoFilterProjector.Project(sheet, autoFilterRange, legacySheet.AutoFilterCriteria);
            }
        }

        private static bool TryGetAutoFilterRange(
            LegacyXlsWorkbook workbook,
            ExcelSheet sheet,
            int worksheetIndex,
            IReadOnlyDictionary<int, int> worksheetIndexByProjectedSheetIndex,
            out string autoFilterRange) {
            autoFilterRange = string.Empty;
            foreach (LegacyXlsDefinedName definedName in workbook.DefinedNames) {
                if (definedName.LocalSheetIndex.HasValue
                    && worksheetIndexByProjectedSheetIndex.TryGetValue(definedName.LocalSheetIndex.Value, out int scopedWorksheetIndex)
                    && scopedWorksheetIndex == worksheetIndex
                    && string.Equals(definedName.Name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                    && TryParseScopedRange(definedName.Reference, sheet, out autoFilterRange)) {
                    return true;
                }
            }

            return false;
        }

        private static IReadOnlyDictionary<int, int> CreateWorksheetIndexByProjectedSheetIndex(LegacyXlsWorkbook workbook) {
            var result = new Dictionary<int, int>();
            int projectedSheetIndex = 0;
            int worksheetIndex = 0;
            foreach (LegacyXlsSheetProjectionEntry sheetEntry in EnumerateSheetsInWorkbookOrder(workbook)) {
                if (sheetEntry.Worksheet != null) {
                    result[projectedSheetIndex] = worksheetIndex++;
                }

                projectedSheetIndex++;
            }

            return result;
        }

        private static void ProjectExternalReferences(LegacyXlsWorkbook workbook, ExcelDocument document) {
            if (!workbook.PreserveExternalWorkbookLinks) {
                return;
            }

            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Workbook workbookRoot = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is null.");

            foreach (LegacyXlsExternalReference reference in workbook.ExternalReferences) {
                if (reference.Kind != LegacyXlsExternalReferenceKind.ExternalWorkbook
                    || string.IsNullOrWhiteSpace(reference.Target)) {
                    continue;
                }

                if (!TryCreateExternalTargetUri(reference.Target!, out Uri? targetUri) || targetUri == null) {
                    continue;
                }

                ExternalWorkbookPart externalWorkbookPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
                ExternalRelationship relationship = externalWorkbookPart.AddExternalRelationship(
                    ExternalLinkPathRelationshipType,
                    targetUri);

                var externalBook = new ExternalBook { Id = relationship.Id };
                if (reference.SheetNames.Count > 0) {
                    externalBook.SheetNames = new SheetNames(reference.SheetNames
                        .Where(sheetName => !string.IsNullOrWhiteSpace(sheetName))
                        .Select(sheetName => new SheetName { Val = sheetName }));
                }

                ExternalDefinedNames? externalDefinedNames = CreateExternalDefinedNames(reference);
                if (externalDefinedNames != null) {
                    externalBook.ExternalDefinedNames = externalDefinedNames;
                }

                externalWorkbookPart.ExternalLink = new ExternalLink(externalBook);
                externalWorkbookPart.ExternalLink.Save();

                ExternalReferences externalReferences = GetOrCreateExternalReferences(workbookRoot);
                externalReferences.Append(new DocumentFormat.OpenXml.Spreadsheet.ExternalReference {
                    Id = workbookPart.GetIdOfPart(externalWorkbookPart)
                });
            }
        }

        private static ExternalReferences GetOrCreateExternalReferences(Workbook workbook) {
            if (workbook.ExternalReferences != null) {
                return workbook.ExternalReferences;
            }

            var externalReferences = new ExternalReferences();
            OpenXmlElement? before = workbook.GetFirstChild<DefinedNames>();
            before ??= workbook.GetFirstChild<CalculationProperties>();
            before ??= workbook.GetFirstChild<OleSize>();
            before ??= workbook.GetFirstChild<CustomWorkbookViews>();
            before ??= workbook.GetFirstChild<PivotCaches>();
            before ??= workbook.GetFirstChild<WebPublishing>();
            before ??= workbook.GetFirstChild<FileRecoveryProperties>();
            before ??= workbook.GetFirstChild<WebPublishObjects>();
            before ??= workbook.GetFirstChild<WorkbookExtensionList>();
            if (before != null) {
                workbook.InsertBefore(externalReferences, before);
            } else {
                workbook.Append(externalReferences);
            }

            return externalReferences;
        }

        private static ExternalDefinedNames? CreateExternalDefinedNames(LegacyXlsExternalReference reference) {
            List<ExternalDefinedName> names = reference.ExternalNames
                .Where(name => !string.IsNullOrWhiteSpace(name.Name))
                .Select(name => {
                    var projectedName = new ExternalDefinedName {
                        Name = name.Name
                    };
                    if (name.LocalSheetIndex.HasValue && name.LocalSheetIndex.Value >= 0) {
                        projectedName.SheetId = (uint)name.LocalSheetIndex.Value;
                    }

                    return projectedName;
                })
                .ToList();

            return names.Count == 0 ? null : new ExternalDefinedNames(names);
        }

        private static bool TryCreateExternalTargetUri(string target, out Uri? uri) {
            string normalized = RemoveControlCharacters(target);
            uri = null;
            if (string.IsNullOrWhiteSpace(normalized)) {
                return false;
            }

            return Uri.TryCreate(normalized, UriKind.RelativeOrAbsolute, out uri);
        }

        private static string RemoveControlCharacters(string value) {
            StringBuilder? sanitized = null;
            for (int i = 0; i < value.Length; i++) {
                if (!char.IsControl(value[i])) {
                    sanitized?.Append(value[i]);
                    continue;
                }

                sanitized ??= new StringBuilder(value.Length).Append(value, 0, i);
            }

            return sanitized == null ? value : sanitized.ToString();
        }

        private static bool TryCreateCommentRichTextRuns(
            LegacyXlsWorkbook workbook,
            LegacyXlsComment comment,
            out IReadOnlyList<ExcelRichTextRun> richTextRuns) {
            return TryCreateRichTextRuns(
                workbook,
                comment.Text,
                comment.FormattingRuns.Select(run => new LegacyXlsTextFormattingRun(run.StartCharacter, run.FontIndex)),
                out richTextRuns);
        }

        private static bool TryCreateCellRichTextRuns(
            LegacyXlsWorkbook workbook,
            string text,
            IReadOnlyList<LegacyXlsTextFormattingRun> formattingRuns,
            out IReadOnlyList<ExcelRichTextRun> richTextRuns) {
            return TryCreateRichTextRuns(workbook, text, formattingRuns, out richTextRuns);
        }

        private static bool TryCreateRichTextRuns(
            LegacyXlsWorkbook workbook,
            string text,
            IEnumerable<LegacyXlsTextFormattingRun> formattingRuns,
            out IReadOnlyList<ExcelRichTextRun> richTextRuns) {
            richTextRuns = Array.Empty<ExcelRichTextRun>();
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            List<LegacyXlsTextFormattingRun> runs = formattingRuns
                .Where(run => run.StartCharacter < text.Length)
                .OrderBy(run => run.StartCharacter)
                .ToList();
            if (runs.Count == 0) {
                return false;
            }

            if (runs[0].StartCharacter != 0) {
                runs.Insert(0, new LegacyXlsTextFormattingRun(0, 0));
            }

            var projectedRuns = new List<ExcelRichTextRun>(runs.Count);
            for (int i = 0; i < runs.Count; i++) {
                int start = runs[i].StartCharacter;
                int end = i + 1 < runs.Count ? runs[i + 1].StartCharacter : text.Length;
                if (end <= start) {
                    continue;
                }

                var projectedRun = new ExcelRichTextRun(text.Substring(start, end - start));
                LegacyXlsFont? font = workbook.GetFont(runs[i].FontIndex);
                if (font != null) {
                    projectedRun.Bold = font.Bold;
                    projectedRun.Italic = font.Italic;
                    projectedRun.Underline = font.Underline;
                    projectedRun.UnderlineStyle = ToUnderlineStyle(font.UnderlineStyle);
                    projectedRun.FontName = font.Name;
                    projectedRun.FontSize = font.Size;
                    projectedRun.VerticalTextAlignment = ToVerticalTextAlignment(font.Escapement);
                    projectedRun.Outline = font.Outline;
                    projectedRun.Shadow = font.Shadow;
                    projectedRun.Condense = font.Condense;
                    projectedRun.Extend = font.Extend;
                    projectedRun.FontFamily = font.Family;
                    projectedRun.FontCharacterSet = font.CharacterSet;
                    if (workbook.TryResolveColor(font.ColorIndex, out string? fontColor)) {
                        projectedRun.FontColor = fontColor;
                    }
                }

                projectedRuns.Add(projectedRun);
            }

            richTextRuns = projectedRuns;
            return projectedRuns.Count > 0;
        }

        private static bool TryParsePrintTitles(
            string reference,
            ExcelSheet scope,
            out int? firstRow,
            out int? lastRow,
            out int? firstColumn,
            out int? lastColumn) {
            firstRow = null;
            lastRow = null;
            firstColumn = null;
            lastColumn = null;

            bool parsedAny = false;
            foreach (string part in SplitDefinedNameReferenceList(reference)) {
                if (!SheetNameLookup.TryParseSheetQualifiedReference(part, out string sheetName, out string localReference, allowExternalWorkbookReferences: false)
                    || !SheetNameLookup.Matches(scope.Name, sheetName)) {
                    return false;
                }

                string normalized = localReference.Replace("$", string.Empty);
                int separator = normalized.IndexOf(':');
                if (separator <= 0 || separator >= normalized.Length - 1) {
                    return false;
                }

                string start = normalized.Substring(0, separator);
                string end = normalized.Substring(separator + 1);
                if (int.TryParse(start, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedFirstRow)
                    && int.TryParse(end, NumberStyles.None, CultureInfo.InvariantCulture, out int parsedLastRow)
                    && parsedFirstRow > 0
                    && parsedLastRow >= parsedFirstRow) {
                    firstRow = parsedFirstRow;
                    lastRow = parsedLastRow;
                    parsedAny = true;
                    continue;
                }

                int parsedFirstColumn = A1.ColumnLettersToIndex(start);
                int parsedLastColumn = A1.ColumnLettersToIndex(end);
                if (parsedFirstColumn <= 0 || parsedLastColumn < parsedFirstColumn) {
                    return false;
                }

                firstColumn = parsedFirstColumn;
                lastColumn = parsedLastColumn;
                parsedAny = true;
            }

            return parsedAny;
        }

        private static bool TryParseScopedRange(string reference, ExcelSheet scope, out string localRange) {
            localRange = string.Empty;
            if (!SheetNameLookup.TryParseSheetQualifiedReference(reference, out string sheetName, out string parsedReference, allowExternalWorkbookReferences: false)
                || !SheetNameLookup.Matches(scope.Name, sheetName)) {
                return false;
            }

            localRange = parsedReference.Replace("$", string.Empty);
            return localRange.Length > 0;
        }

        private static IReadOnlyList<string> SplitDefinedNameReferenceList(string text) {
            var parts = new List<string>();
            var current = new System.Text.StringBuilder(text.Length);
            bool inQuote = false;

            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                if (ch == '\'') {
                    current.Append(ch);
                    if (inQuote && i + 1 < text.Length && text[i + 1] == '\'') {
                        current.Append(text[++i]);
                    } else {
                        inQuote = !inQuote;
                    }
                    continue;
                }

                if (ch == ',' && !inQuote) {
                    string part = current.ToString().Trim();
                    if (part.Length > 0) {
                        parts.Add(part);
                    }

                    current.Clear();
                    continue;
                }

                current.Append(ch);
            }

            string finalPart = current.ToString().Trim();
            if (finalPart.Length > 0) {
                parts.Add(finalPart);
            }

            return parts;
        }

        private static void ProjectPageBreaks(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            foreach (LegacyXlsPageBreak pageBreak in legacySheet.RowPageBreaks) {
                sheet.AddManualRowPageBreak(pageBreak.Position, save: false);
            }

            foreach (LegacyXlsPageBreak pageBreak in legacySheet.ColumnPageBreaks) {
                sheet.AddManualColumnPageBreak(pageBreak.Position, save: false);
            }
        }

        private static void ProjectPageSetup(LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            LegacyXlsPageSetup? pageSetup = legacySheet.PageSetup;
            if (pageSetup == null) {
                return;
            }

            if (pageSetup.LeftMargin.HasValue
                || pageSetup.RightMargin.HasValue
                || pageSetup.TopMargin.HasValue
                || pageSetup.BottomMargin.HasValue
                || pageSetup.HeaderMargin.HasValue
                || pageSetup.FooterMargin.HasValue) {
                sheet.SetMargins(
                    pageSetup.LeftMargin ?? 0.7d,
                    pageSetup.RightMargin ?? 0.7d,
                    pageSetup.TopMargin ?? 0.75d,
                    pageSetup.BottomMargin ?? 0.75d,
                    pageSetup.HeaderMargin ?? 0.3d,
                    pageSetup.FooterMargin ?? 0.3d);
            }

            if (pageSetup.Landscape.HasValue) {
                sheet.SetOrientation(pageSetup.Landscape.Value ? ExcelPageOrientation.Landscape : ExcelPageOrientation.Portrait);
            }

            if (pageSetup.PrintGridLines.HasValue
                || pageSetup.PrintHeadings.HasValue
                || pageSetup.HorizontalCentered.HasValue
                || pageSetup.VerticalCentered.HasValue) {
                sheet.SetPrintOptions(
                    pageSetup.PrintGridLines,
                    pageSetup.PrintHeadings,
                    pageSetup.HorizontalCentered,
                    pageSetup.VerticalCentered,
                    save: false);
            }

            uint? fitToWidth = pageSetup.FitToWidth.HasValue ? pageSetup.FitToWidth.Value : null;
            uint? fitToHeight = pageSetup.FitToHeight.HasValue ? pageSetup.FitToHeight.Value : null;
            uint? scale = pageSetup.Scale.HasValue ? pageSetup.Scale.Value : null;
            if (fitToWidth.HasValue || fitToHeight.HasValue || scale.HasValue || pageSetup.PageOrder.HasValue) {
                sheet.SetPageSetup(fitToWidth, fitToHeight, scale, pageSetup.PageOrder);
            }

            if (pageSetup.FitToPage.HasValue) {
                sheet.SetFitToPage(pageSetup.FitToPage.Value);
            }

            if (!string.IsNullOrEmpty(pageSetup.HeaderText)
                || !string.IsNullOrEmpty(pageSetup.FooterText)
                || pageSetup.DifferentFirstHeaderFooter.HasValue
                || pageSetup.DifferentOddEvenHeaderFooter.HasValue
                || pageSetup.ScaleHeaderFooterWithDocument.HasValue
                || pageSetup.AlignHeaderFooterWithMargins.HasValue) {
                var header = SplitHeaderFooterText(pageSetup.HeaderText);
                var footer = SplitHeaderFooterText(pageSetup.FooterText);
                sheet.SetHeaderFooter(
                    header.Left,
                    header.Center,
                    header.Right,
                    footer.Left,
                    footer.Center,
                    footer.Right,
                    pageSetup.DifferentFirstHeaderFooter == true,
                    pageSetup.DifferentOddEvenHeaderFooter == true,
                    pageSetup.AlignHeaderFooterWithMargins ?? true,
                    pageSetup.ScaleHeaderFooterWithDocument ?? true);
            }

            if (pageSetup.DifferentFirstHeaderFooter == true
                || !string.IsNullOrEmpty(pageSetup.FirstHeaderText)
                || !string.IsNullOrEmpty(pageSetup.FirstFooterText)) {
                var firstHeader = SplitHeaderFooterText(pageSetup.FirstHeaderText);
                var firstFooter = SplitHeaderFooterText(pageSetup.FirstFooterText);
                sheet.SetFirstPageHeaderFooter(
                    firstHeader.Left,
                    firstHeader.Center,
                    firstHeader.Right,
                    firstFooter.Left,
                    firstFooter.Center,
                    firstFooter.Right,
                    enabled: true);
            }

            if (pageSetup.DifferentOddEvenHeaderFooter == true
                || !string.IsNullOrEmpty(pageSetup.EvenHeaderText)
                || !string.IsNullOrEmpty(pageSetup.EvenFooterText)) {
                var evenHeader = SplitHeaderFooterText(pageSetup.EvenHeaderText);
                var evenFooter = SplitHeaderFooterText(pageSetup.EvenFooterText);
                sheet.SetEvenPageHeaderFooter(
                    evenHeader.Left,
                    evenHeader.Center,
                    evenHeader.Right,
                    evenFooter.Left,
                    evenFooter.Center,
                    evenFooter.Right,
                    enabled: true);
            }
        }

        private static (string? Left, string? Center, string? Right) SplitHeaderFooterText(string? text) {
            if (string.IsNullOrEmpty(text)) {
                return (null, null, null);
            }

            if (!ContainsHeaderFooterSectionMarker(text!)) {
                return (null, text, null);
            }

            string? left = null;
            string? center = null;
            string? right = null;
            int i = 0;
            while (i < text!.Length) {
                if (text[i] != '&' || i + 1 >= text.Length) {
                    i++;
                    continue;
                }

                char section = text[i + 1];
                if (section != 'L' && section != 'C' && section != 'R') {
                    i += 2;
                    continue;
                }

                i += 2;
                int start = i;
                while (i < text.Length) {
                    if (text[i] == '&' && i + 1 < text.Length) {
                        char next = text[i + 1];
                        if (next == 'L' || next == 'C' || next == 'R') {
                            break;
                        }
                    }

                    i++;
                }

                string value = text.Substring(start, i - start);
                if (section == 'L') {
                    left = AppendHeaderFooterSection(left, value);
                } else if (section == 'C') {
                    center = AppendHeaderFooterSection(center, value);
                } else {
                    right = AppendHeaderFooterSection(right, value);
                }
            }

            return (left, center, right);
        }

        private static bool ContainsHeaderFooterSectionMarker(string text) {
            return text.IndexOf("&L", StringComparison.Ordinal) >= 0
                || text.IndexOf("&C", StringComparison.Ordinal) >= 0
                || text.IndexOf("&R", StringComparison.Ordinal) >= 0;
        }

        private static string? AppendHeaderFooterSection(string? current, string value) {
            if (value.Length == 0) {
                return current;
            }

            return current == null ? value : current + value;
        }

        private static object? GetProjectedCellValue(
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (cell.Kind != LegacyXlsCellValueKind.Number || format?.IsDateLike != true || cell.Value is not double serial) {
                return cell.Value;
            }

            return LegacyXlsDateSerialConverter.TryConvert(serial, workbook.Uses1904DateSystem, out DateTime value)
                ? value
                : cell.Value;
        }

        private static void ApplyNumberFormat(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.NumberFormatCode == null || format.NumberFormatId == 0) {
                return;
            }

            if (format.IsBuiltInNumberFormat) {
                sheet.FormatCellBuiltInNumberFormat(cell.Row, cell.Column, format.NumberFormatId);
                return;
            }

            sheet.FormatCell(cell.Row, cell.Column, format.NumberFormatCode);
        }

        private static void ApplyCellFormat(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format == null) {
                return;
            }

            uint styleIndex = sheet.GetOrCreateLegacyCellFormatStyleIndex(workbook, format);
            if (styleIndex != 0U) {
                sheet.SetCellStyleIndex(cell.Row, cell.Column, styleIndex, save: false);
            }
        }

        private static void ApplyFont(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format == null) {
                return;
            }

            LegacyXlsFont? font = workbook.GetFont(format.FontIndex);
            if (font == null) {
                return;
            }

            workbook.TryResolveColor(font.ColorIndex, out string? fontColor);
            VerticalAlignmentRunValues? verticalTextAlignment = ToVerticalTextAlignment(font.Escapement);
            UnderlineValues? underlineStyle = ToUnderlineStyle(font.UnderlineStyle);
            if (font.Name == null && !font.Size.HasValue && fontColor == null && !font.Bold && !font.Italic && !font.Underline && !font.Strikeout && !verticalTextAlignment.HasValue && !font.Outline && !font.Shadow && !font.Condense && !font.Extend) {
                return;
            }

            sheet.FormatCellFont(
                cell.Row,
                cell.Column,
                font.Name,
                font.Size,
                fontColor,
                font.Bold,
                font.Italic,
                font.Underline,
                font.Strikeout,
                underlineStyle,
                verticalTextAlignment,
                font.Family,
                font.CharacterSet,
                font.Outline,
                font.Shadow,
                font.Condense,
                font.Extend);
        }

        private static VerticalAlignmentRunValues? ToVerticalTextAlignment(LegacyXlsFontEscapement escapement) {
            return escapement == LegacyXlsFontEscapement.Superscript
                ? VerticalAlignmentRunValues.Superscript
                : escapement == LegacyXlsFontEscapement.Subscript
                    ? VerticalAlignmentRunValues.Subscript
                    : null;
        }

        private static UnderlineValues? ToUnderlineStyle(byte underlineStyle) {
            return underlineStyle switch {
                0x01 => UnderlineValues.Single,
                0x02 => UnderlineValues.Double,
                0x21 => UnderlineValues.SingleAccounting,
                0x22 => UnderlineValues.DoubleAccounting,
                _ => null
            };
        }

        private static void ApplyFill(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format == null || !format.ApplyFill || format.FillPattern == 0) {
                return;
            }

            PatternValues? pattern = ToFillPattern(format.FillPattern);
            if (!pattern.HasValue) {
                return;
            }

            string? foregroundColor = ResolveColor(workbook, format.FillForegroundColorIndex);
            string? backgroundColor = ResolveColor(workbook, format.FillBackgroundColorIndex);
            if (foregroundColor == null && backgroundColor == null) {
                return;
            }

            if (format.FillPattern == 1 && foregroundColor != null) {
                sheet.FormatCellFill(cell.Row, cell.Column, PatternValues.Solid, foregroundColor, foregroundColor);
                return;
            }

            sheet.FormatCellFill(cell.Row, cell.Column, pattern.Value, foregroundColor, backgroundColor);
        }

        private static void ApplyAlignment(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.ApplyAlignment != true) {
                return;
            }

            sheet.FormatCellAlignment(
                cell.Row,
                cell.Column,
                ToHorizontalAlignment(format.HorizontalAlignment),
                ToVerticalAlignment(format.VerticalAlignment),
                format.WrapText,
                ToTextRotation(format.TextRotation),
                format.Indent,
                format.ShrinkToFit,
                ToReadingOrder(format.ReadingOrder));
        }

        private static HorizontalAlignmentValues? ToHorizontalAlignment(byte alignment) {
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

        private static VerticalAlignmentValues? ToVerticalAlignment(byte alignment) {
            return alignment switch {
                0 => VerticalAlignmentValues.Top,
                1 => VerticalAlignmentValues.Center,
                2 => VerticalAlignmentValues.Bottom,
                3 => VerticalAlignmentValues.Justify,
                4 => VerticalAlignmentValues.Distributed,
                _ => null
            };
        }

        private static uint? ToTextRotation(byte rotation) {
            return rotation <= 180 || rotation == 255 ? rotation : null;
        }

        private static uint? ToReadingOrder(byte readingOrder) {
            return readingOrder <= 2 ? readingOrder : null;
        }

        private static void ApplyBorder(
            ExcelSheet sheet,
            LegacyXlsWorkbook workbook,
            LegacyXlsCell cell,
            LegacyXlsCellFormat? format) {
            if (format?.Border == null) {
                return;
            }

            LegacyXlsBorder border = format.Border;
            sheet.FormatCellBorder(
                cell.Row,
                cell.Column,
                ToBorderStyle(border.LeftStyle),
                ResolveColor(workbook, border.LeftColorIndex),
                ToBorderStyle(border.RightStyle),
                ResolveColor(workbook, border.RightColorIndex),
                ToBorderStyle(border.TopStyle),
                ResolveColor(workbook, border.TopColorIndex),
                ToBorderStyle(border.BottomStyle),
                ResolveColor(workbook, border.BottomColorIndex),
                ToBorderStyle(border.DiagonalStyle),
                ResolveColor(workbook, border.DiagonalColorIndex),
                border.DiagonalUp,
                border.DiagonalDown);
        }

        private static void ApplyProtection(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.ApplyProtection != true) {
                return;
            }

            sheet.FormatCellProtection(cell.Row, cell.Column, format.Locked, format.FormulaHidden);
        }

        private static void ApplyQuotePrefix(ExcelSheet sheet, LegacyXlsCell cell, LegacyXlsCellFormat? format) {
            if (format?.QuotePrefix != true) {
                return;
            }

            sheet.FormatCellQuotePrefix(cell.Row, cell.Column, true);
        }

        private static BorderStyleValues? ToBorderStyle(byte style) {
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

        private static PatternValues? ToFillPattern(byte pattern) {
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

        private static string? ResolveColor(LegacyXlsWorkbook workbook, ushort colorIndex) {
            return workbook.TryResolveColor(colorIndex, out string? color) ? color : null;
        }

        private static string ToA1Range(LegacyXlsMergedRange mergedRange) {
            string start = A1.CellReference(mergedRange.StartRow, mergedRange.StartColumn);
            string end = A1.CellReference(mergedRange.EndRow, mergedRange.EndColumn);
            return start == end ? start : start + ":" + end;
        }

        private static string ToA1Range(LegacyXlsHyperlink hyperlink) {
            string start = A1.CellReference(hyperlink.StartRow, hyperlink.StartColumn);
            string end = A1.CellReference(hyperlink.EndRow, hyperlink.EndColumn);
            return start == end ? start : start + ":" + end;
        }

    }
}
