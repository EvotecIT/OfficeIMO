using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Projection;

/// <summary>Keeps projected sheet metadata attached to native tab identities through structure edits.</summary>
internal static partial class LegacyXlsWorkbookProjector {
        private static void ProjectCodeNames(
            LegacyXlsWorkbook workbook,
            ExcelDocument document,
            CancellationToken cancellationToken) {
            if (!string.IsNullOrWhiteSpace(workbook.CodeName)) {
                WorkbookProperties properties = GetOrCreateWorkbookProperties(document);
                properties.CodeName = workbook.CodeName;
            }

            var targets = document.Sheets.ToDictionary(sheet => sheet.Name, StringComparer.OrdinalIgnoreCase);
            foreach (var source in workbook.Worksheets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!targets.TryGetValue(source.Name, out var target)) continue;
                string? codeName = source.CodeName;
                if (string.IsNullOrWhiteSpace(codeName)) {
                    continue;
                }

                Worksheet worksheet = target.WorksheetPart.Worksheet!;
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

        private static void ProjectWorkbookWindow(
            LegacyXlsWorkbook workbook,
            ExcelDocument document,
            CancellationToken cancellationToken) {
            if (workbook.Windows.Count == 0) {
                return;
            }

            LegacyXlsWorkbookWindow window = workbook.Windows[0];
            var ordered = EnumerateSheetsInWorkbookOrder(workbook, cancellationToken);
            var projectedTabByNativeIndex = new Dictionary<int, int>();
            for (int index = 0; index < ordered.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (workbook.NativeSheetOrder.TryGetValue(ordered[index].StreamOffset, out int nativeIndex) &&
                    !projectedTabByNativeIndex.ContainsKey(nativeIndex))
                    projectedTabByNativeIndex.Add(nativeIndex, index);
            }
            int active = MapNativeTab(projectedTabByNativeIndex, window.ActiveSheetIndex);
            if (active < ordered.Count && ordered[active].Worksheet is { Visibility: 0 } activeSheet) document.SetActiveWorksheet(activeSheet.Name);

            BookViews workbookViews = GetOrCreateWorkbookViews(document);
            workbookViews.RemoveAllChildren<WorkbookView>();
            foreach (LegacyXlsWorkbookWindow legacyWindow in workbook.Windows) {
                cancellationToken.ThrowIfCancellationRequested();
                WorkbookView workbookView = workbookViews.AppendChild(new WorkbookView());
                ProjectWorkbookWindowView(legacyWindow, workbookView, projectedTabByNativeIndex);
            }
        }

        private static void ProjectWorkbookWindowView(LegacyXlsWorkbookWindow window, WorkbookView workbookView,
            IReadOnlyDictionary<int, int> projectedTabByNativeIndex) {
            workbookView.XWindow = window.HorizontalPositionTwips;
            workbookView.YWindow = window.VerticalPositionTwips;
            workbookView.WindowWidth = checked((uint)Math.Max(0, (int)window.WidthTwips));
            workbookView.WindowHeight = checked((uint)Math.Max(0, (int)window.HeightTwips));
            workbookView.Visibility = ToWorkbookViewVisibility(window);
            workbookView.Minimized = window.Minimized;
            workbookView.ActiveTab = checked((uint)MapNativeTab(projectedTabByNativeIndex, window.ActiveSheetIndex));
            workbookView.FirstSheet = checked((uint)MapNativeTab(projectedTabByNativeIndex, window.FirstVisibleSheetTabIndex));
            workbookView.ShowHorizontalScroll = window.HorizontalScrollBarVisible;
            workbookView.ShowVerticalScroll = window.VerticalScrollBarVisible;
            workbookView.ShowSheetTabs = window.SheetTabsVisible;
            workbookView.TabRatio = window.SheetTabRatio;
        }

        private static int MapNativeTab(IReadOnlyDictionary<int, int> projectedTabByNativeIndex, int nativeIndex) =>
            projectedTabByNativeIndex.TryGetValue(nativeIndex, out int projectedIndex) ? projectedIndex : 0;

        private static void ProjectAutoFilters(
            LegacyXlsWorkbook workbook,
            ExcelDocument document,
            CancellationToken cancellationToken) {
            IReadOnlyDictionary<int, int> worksheetIndexByProjectedSheetIndex =
                CreateWorksheetIndexByProjectedSheetIndex(workbook, cancellationToken);
            var targets = document.Sheets.Select((sheet, index) => (sheet, index)).ToDictionary(item => item.sheet.Name, StringComparer.OrdinalIgnoreCase);
            foreach (var legacySheet in workbook.Worksheets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!targets.TryGetValue(legacySheet.Name, out var target)) continue;
                if (legacySheet.AutoFilterCriteria.Count == 0) {
                    continue;
                }

                ExcelSheet sheet = target.sheet;
                if (!TryGetAutoFilterRange(
                    workbook,
                    sheet,
                    target.index,
                    worksheetIndexByProjectedSheetIndex,
                    cancellationToken,
                    out string autoFilterRange)) {
                    continue;
                }

                LegacyXlsAutoFilterProjector.Project(sheet, autoFilterRange, legacySheet.AutoFilterCriteria);
            }
        }

}
