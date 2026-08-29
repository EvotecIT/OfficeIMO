using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X14SlicerDrawing = DocumentFormat.OpenXml.Office2010.Drawing.Slicer;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X15TimelineDrawing = DocumentFormat.OpenXml.Office2013.Drawing.TimeSlicer;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string SlicerCachesExtensionUri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";
        private const string SlicerListExtensionUri = "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}";
        private const string TimelineCachesExtensionUri = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}";
        private const string TimelineListExtensionUri = "{7E03D99C-DC04-49D9-9315-930204A7B6E9}";
        private const string PivotSlicerExtensionUri = "{725AE2AE-9491-48BE-B2B4-4EB974FC3084}";
        private const string PivotTimelineExtensionUri = "{03082B11-2C62-411C-B77F-237D8FCFBE4C}";
        private const string SlicerGraphicDataUri = "http://schemas.microsoft.com/office/drawing/2010/slicer";
        private const string TimelineGraphicDataUri = "http://schemas.microsoft.com/office/drawing/2012/timeslicer";

        /// <summary>Adds a native Excel slicer view bound to a PivotTable field.</summary>
        public ExcelPivotInteractionInfo AddPivotSlicer(
            string pivotTableName,
            string sourceField,
            string worksheetName,
            ExcelSlicerViewOptions? options = null) {
            options ??= new ExcelSlicerViewOptions();
            ExcelPivotTableInfo pivot = ValidatePivotInteractionBinding(pivotTableName, sourceField);
            ExcelSheet targetSheet = ResolvePivotInteractionWorksheet(worksheetName);
            ValidatePivotInteractionPlacement(options.Row, options.Column, options.WidthPixels, options.HeightPixels);
            if (options.ItemColumns < 1 || options.ItemColumns > 20000) {
                throw new ArgumentOutOfRangeException(nameof(options), "Slicer item columns must be between 1 and 20,000.");
            }

            ExcelPivotInteractionInfo? created = null;
            targetSheet.ApplyTransactionalMutation(_ => {
                string viewName = ResolveNativeInteractionName(
                    options.Name,
                    "Slicer_" + sourceField.Trim(),
                    EnumeratePivotInteractionNames());
                PivotTablePart pivotPart = FindPivotTablePart(pivot)
                    ?? throw new InvalidOperationException($"Pivot table '{pivot.Name}' has no package part.");
                SlicerCachePart cachePart = ResolveOrCreateSlicerCache(
                    pivot,
                    pivotPart,
                    sourceField.Trim(),
                    options.CacheName);
                string cacheName = cachePart.SlicerCacheDefinition?.Name?.Value
                    ?? throw new InvalidDataException("Native slicer cache has no name.");

                SlicersPart slicersPart = targetSheet.WorksheetPart.SlicersParts.FirstOrDefault()
                    ?? targetSheet.WorksheetPart.AddNewPart<SlicersPart>();
                slicersPart.Slicers ??= new X14.Slicers();
                slicersPart.Slicers!.Append(new X14.Slicer {
                    Name = viewName,
                    Cache = cacheName,
                    Caption = string.IsNullOrWhiteSpace(options.Caption) ? sourceField.Trim() : options.Caption!.Trim(),
                    ColumnCount = (uint)options.ItemColumns,
                    ShowCaption = options.ShowCaption,
                    LockedPosition = options.LockedPosition,
                    RowHeight = 19U * 12700U,
                    Style = ValidateBuiltInInteractionStyle(options.Style, "SlicerStyle", 6)
                });
                slicersPart.Slicers!.Save();
                EnsureWorksheetSlicerReference(targetSheet.WorksheetPart, slicersPart);
                AddPivotInteractionDrawing(
                    targetSheet,
                    viewName,
                    options.Row,
                    options.Column,
                    options.WidthPixels,
                    options.HeightPixels,
                    timeline: false);
                targetSheet.WorksheetPart.Worksheet!.Save();
                MarkMetadataPartChanged();
                created = new ExcelPivotInteractionInfo(
                    ExcelPivotInteractionCacheKind.Slicer,
                    viewName,
                    cacheName,
                    sourceField.Trim(),
                    pivot.Name,
                    targetSheet.Name,
                    targetSheet.WorksheetPart.GetIdOfPart(slicersPart));
                return 1;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
            return created!;
        }

        /// <summary>Adds a native Excel timeline view bound to a date-only PivotTable field.</summary>
        public ExcelPivotInteractionInfo AddPivotTimeline(
            string pivotTableName,
            string sourceField,
            string worksheetName,
            ExcelTimelineViewOptions? options = null) {
            options ??= new ExcelTimelineViewOptions();
            ExcelSheet targetSheet = ResolvePivotInteractionWorksheet(worksheetName);
            ValidatePivotInteractionPlacement(options.Row, options.Column, options.WidthPixels, options.HeightPixels);
            if (!Enum.IsDefined(typeof(ExcelTimelineLevel), options.Level)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Timeline level is invalid.");
            }

            ExcelPivotInteractionInfo? created = null;
            targetSheet.ApplyTransactionalMutation(_ => {
                ExcelPivotTableInfo pivot = ValidatePivotInteractionBinding(pivotTableName, sourceField);
                PivotTablePart pivotPart = FindPivotTablePart(pivot)
                    ?? throw new InvalidOperationException($"Pivot table '{pivot.Name}' has no package part.");
                if (!IsDateOnlyPivotSourceField(pivot, sourceField.Trim())) {
                    throw new ArgumentException(
                        $"Field '{sourceField}' is not a date-only source field and cannot be used for a timeline.",
                        nameof(sourceField));
                }
                TimeLineCachePart? compatibleCache = FindCompatibleTimelineCache(
                    pivot,
                    sourceField.Trim(),
                    options.CacheName);
                (DateTime Minimum, DateTime Maximum)? sourceBounds = null;
                if (compatibleCache == null) {
                    CacheField sourceCacheField = GetPivotSourceCacheField(pivotPart, sourceField.Trim(), out int _);
                    sourceBounds = GetTimelineBounds(sourceCacheField, pivot);
                }
                string viewName = ResolveNativeInteractionName(
                    options.Name,
                    "Timeline_" + sourceField.Trim(),
                    EnumeratePivotInteractionNames());
                TimeLineCachePart cachePart = ResolveOrCreateTimelineCache(
                    pivot,
                    pivotPart,
                    sourceField.Trim(),
                    options.CacheName,
                    compatibleCache,
                    sourceBounds);
                string cacheName = cachePart.TimelineCacheDefinition?.Name?.Value
                    ?? throw new InvalidDataException("Native timeline cache has no name.");

                TimeLinePart timelinesPart = targetSheet.WorksheetPart.TimeLineParts.FirstOrDefault()
                    ?? targetSheet.WorksheetPart.AddNewPart<TimeLinePart>();
                timelinesPart.Timelines ??= new X15.Timelines();
                timelinesPart.Timelines!.Append(new X15.Timeline {
                    Name = viewName,
                    Cache = cacheName,
                    Caption = string.IsNullOrWhiteSpace(options.Caption) ? sourceField.Trim() : options.Caption!.Trim(),
                    ShowHeader = options.ShowHeader,
                    ShowSelectionLabel = options.ShowSelectionLabel,
                    ShowTimeLevel = options.ShowTimeLevel,
                    ShowHorizontalScrollbar = options.ShowHorizontalScrollbar,
                    Level = (uint)options.Level,
                    SelectionLevel = (uint)options.Level,
                    Style = ValidateBuiltInInteractionStyle(options.Style, "TimelineStyle", 6)
                });
                timelinesPart.Timelines!.Save();
                EnsureWorksheetTimelineReference(targetSheet.WorksheetPart, timelinesPart);
                AddPivotInteractionDrawing(
                    targetSheet,
                    viewName,
                    options.Row,
                    options.Column,
                    options.WidthPixels,
                    options.HeightPixels,
                    timeline: true);
                targetSheet.WorksheetPart.Worksheet!.Save();
                MarkMetadataPartChanged();
                created = new ExcelPivotInteractionInfo(
                    ExcelPivotInteractionCacheKind.Timeline,
                    viewName,
                    cacheName,
                    sourceField.Trim(),
                    pivot.Name,
                    targetSheet.Name,
                    targetSheet.WorksheetPart.GetIdOfPart(timelinesPart));
                return 1;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
            return created!;
        }

        /// <summary>Returns native slicer and timeline views across the workbook.</summary>
        public IReadOnlyList<ExcelPivotInteractionInfo> GetPivotInteractions() {
            return ExecuteReadAfterMaterializing(() => {
                var interactions = new List<ExcelPivotInteractionInfo>();
                var slicerCaches = WorkbookPartRoot.SlicerCacheParts
                    .Where(part => part.SlicerCacheDefinition != null)
                    .GroupBy(part => part.SlicerCacheDefinition!.Name?.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
                var timelineCaches = WorkbookPartRoot.TimeLineCacheParts
                    .Where(part => part.TimelineCacheDefinition != null)
                    .GroupBy(part => part.TimelineCacheDefinition!.Name?.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

                foreach (ExcelSheet sheet in GetSheetsForLockedOperation()) {
                    foreach (SlicersPart part in sheet.WorksheetPart.SlicersParts) {
                        foreach (X14.Slicer view in part.Slicers?.Elements<X14.Slicer>() ?? Enumerable.Empty<X14.Slicer>()) {
                            string cacheName = view.Cache?.Value ?? string.Empty;
                            slicerCaches.TryGetValue(cacheName, out SlicerCachePart? cache);
                            X14.SlicerCachePivotTable? pivot = cache?.SlicerCacheDefinition?
                                .SlicerCachePivotTables?.Elements<X14.SlicerCachePivotTable>().FirstOrDefault();
                            interactions.Add(new ExcelPivotInteractionInfo(
                                ExcelPivotInteractionCacheKind.Slicer,
                                view.Name?.Value ?? string.Empty,
                                cacheName,
                                cache?.SlicerCacheDefinition?.SourceName?.Value ?? string.Empty,
                                pivot?.Name?.Value,
                                sheet.Name,
                                sheet.WorksheetPart.GetIdOfPart(part)));
                        }
                    }
                    foreach (TimeLinePart part in sheet.WorksheetPart.TimeLineParts) {
                        foreach (X15.Timeline view in part.Timelines?.Elements<X15.Timeline>() ?? Enumerable.Empty<X15.Timeline>()) {
                            string cacheName = view.Cache?.Value ?? string.Empty;
                            timelineCaches.TryGetValue(cacheName, out TimeLineCachePart? cache);
                            X15.TimelineCachePivotTable? pivot = cache?.TimelineCacheDefinition?
                                .TimelineCachePivotTables?.Elements<X15.TimelineCachePivotTable>().FirstOrDefault();
                            interactions.Add(new ExcelPivotInteractionInfo(
                                ExcelPivotInteractionCacheKind.Timeline,
                                view.Name?.Value ?? string.Empty,
                                cacheName,
                                cache?.TimelineCacheDefinition?.SourceName?.Value ?? string.Empty,
                                pivot?.Name?.Value,
                                sheet.Name,
                                sheet.WorksheetPart.GetIdOfPart(part)));
                        }
                    }
                }
                return interactions
                    .OrderBy(item => item.WorksheetName, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            });
        }

        /// <summary>Removes one native slicer or timeline view and optionally its unused cache.</summary>
        public bool RemovePivotInteraction(string name, bool removeUnusedCache = true) {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentNullException(nameof(name));
            ExcelPivotInteractionInfo? interaction = GetPivotInteractions().FirstOrDefault(item =>
                string.Equals(item.Name, name.Trim(), StringComparison.OrdinalIgnoreCase));
            if (interaction == null) return false;
            ExcelSheet sheet = this[interaction.WorksheetName];
            bool removed = false;
            sheet.ApplyTransactionalMutation(_ => {
                if (interaction.Kind == ExcelPivotInteractionCacheKind.Slicer) {
                    SlicersPart? part = sheet.WorksheetPart.SlicersParts.FirstOrDefault(candidate =>
                        candidate.Slicers?.Elements<X14.Slicer>().Any(view =>
                            string.Equals(view.Name?.Value, interaction.Name, StringComparison.OrdinalIgnoreCase)) == true);
                    X14.Slicer? view = part?.Slicers?.Elements<X14.Slicer>().FirstOrDefault(candidate =>
                        string.Equals(candidate.Name?.Value, interaction.Name, StringComparison.OrdinalIgnoreCase));
                    if (view != null) {
                        view.Remove();
                        RemovePivotInteractionDrawing(sheet, interaction.Name, timeline: false);
                        CleanupSlicerViewPart(sheet.WorksheetPart, part!);
                        removed = true;
                    }
                } else {
                    TimeLinePart? part = sheet.WorksheetPart.TimeLineParts.FirstOrDefault(candidate =>
                        candidate.Timelines?.Elements<X15.Timeline>().Any(view =>
                            string.Equals(view.Name?.Value, interaction.Name, StringComparison.OrdinalIgnoreCase)) == true);
                    X15.Timeline? view = part?.Timelines?.Elements<X15.Timeline>().FirstOrDefault(candidate =>
                        string.Equals(candidate.Name?.Value, interaction.Name, StringComparison.OrdinalIgnoreCase));
                    if (view != null) {
                        view.Remove();
                        RemovePivotInteractionDrawing(sheet, interaction.Name, timeline: true);
                        CleanupTimelineViewPart(sheet.WorksheetPart, part!);
                        removed = true;
                    }
                }

                if (removed
                    && removeUnusedCache
                    && !IsNativeInteractionCacheInUse(interaction.Kind, interaction.CacheName)) {
                    RemoveNativeInteractionCache(interaction.Kind, interaction.CacheName);
                }
                if (removed) MarkMetadataPartChanged();
                return removed ? 1 : 0;
            }, new ExcelMutationPlanOptions(), CancellationToken.None);
            return removed;
        }

        private bool IsNativeInteractionCacheInUse(
            ExcelPivotInteractionCacheKind kind,
            string cacheName) {
            if (kind == ExcelPivotInteractionCacheKind.Slicer) {
                return WorkbookPartRoot.WorksheetParts
                    .SelectMany(part => part.SlicersParts)
                    .SelectMany(part => part.Slicers?.Elements<X14.Slicer>() ?? Enumerable.Empty<X14.Slicer>())
                    .Any(view => string.Equals(
                        view.Cache?.Value,
                        cacheName,
                        StringComparison.OrdinalIgnoreCase));
            }

            return WorkbookPartRoot.WorksheetParts
                .SelectMany(part => part.TimeLineParts)
                .SelectMany(part => part.Timelines?.Elements<X15.Timeline>() ?? Enumerable.Empty<X15.Timeline>())
                .Any(view => string.Equals(
                    view.Cache?.Value,
                    cacheName,
                    StringComparison.OrdinalIgnoreCase));
        }

        private ExcelSheet ResolvePivotInteractionWorksheet(string worksheetName) {
            if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentNullException(nameof(worksheetName));
            return Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, worksheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException($"Worksheet '{worksheetName}' was not found.", nameof(worksheetName));
        }

        private static void ValidatePivotInteractionPlacement(int row, int column, int widthPixels, int heightPixels) {
            if (row < 1 || row > 1_048_576) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > 16_384) throw new ArgumentOutOfRangeException(nameof(column));
            if (widthPixels < 1 || heightPixels < 1) throw new ArgumentOutOfRangeException(nameof(widthPixels));
        }

        private static string ValidateBuiltInInteractionStyle(string style, string prefix, int maximum) {
            if (string.IsNullOrWhiteSpace(style)) return prefix + "Light2";
            string normalized = style.Trim();
            bool valid = Enumerable.Range(1, maximum).Any(index =>
                string.Equals(normalized, prefix + "Light" + index.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
                || string.Equals(normalized, prefix + "Dark" + index.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal));
            if (!valid && prefix == "SlicerStyle") {
                valid = normalized == "SlicerStyleOther1" || normalized == "SlicerStyleOther2";
            }
            return valid ? normalized : throw new ArgumentException($"'{style}' is not a supported built-in {prefix}.", nameof(style));
        }

        private static string ResolveNativeInteractionName(string? requested, string fallback, IEnumerable<string> existing) {
            var names = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
            string candidate = string.IsNullOrWhiteSpace(requested)
                ? CreatePivotInteractionCacheName(string.Empty, fallback).TrimStart('_')
                : requested!.Trim();
            if (candidate.Length == 0 || candidate.Length > 255) {
                throw new ArgumentException("Interaction names must contain 1 to 255 characters.", nameof(requested));
            }
            if (!names.Contains(candidate)) return candidate;
            if (!string.IsNullOrWhiteSpace(requested)) {
                throw new InvalidOperationException($"Pivot interaction '{candidate}' already exists.");
            }
            for (int suffix = 2; suffix <= names.Count + 1; suffix++) {
                string suffixText = "_" + suffix.ToString(CultureInfo.InvariantCulture);
                string generated = candidate.Substring(0, Math.Min(candidate.Length, 255 - suffixText.Length)) + suffixText;
                if (!names.Contains(generated)) return generated;
            }
            throw new InvalidOperationException("Unable to allocate a unique pivot interaction name.");
        }

        private IEnumerable<string> EnumeratePivotInteractionNames() {
            foreach (WorksheetPart worksheetPart in WorkbookPartRoot.WorksheetParts) {
                foreach (SlicersPart part in worksheetPart.SlicersParts) {
                    foreach (X14.Slicer view in part.Slicers?.Elements<X14.Slicer>() ?? Enumerable.Empty<X14.Slicer>()) {
                        if (!string.IsNullOrWhiteSpace(view.Name?.Value)) yield return view.Name!.Value!;
                    }
                }
                foreach (TimeLinePart part in worksheetPart.TimeLineParts) {
                    foreach (X15.Timeline view in part.Timelines?.Elements<X15.Timeline>() ?? Enumerable.Empty<X15.Timeline>()) {
                        if (!string.IsNullOrWhiteSpace(view.Name?.Value)) yield return view.Name!.Value!;
                    }
                }
            }
        }

        private SlicerCachePart ResolveOrCreateSlicerCache(
            ExcelPivotTableInfo pivot,
            PivotTablePart pivotPart,
            string sourceField,
            string? requestedCacheName) {
            SlicerCachePart? existing = FindCompatibleSlicerCache(pivot, sourceField, requestedCacheName);
            if (existing != null) {
                X14.SlicerCachePivotTables targets = existing.SlicerCacheDefinition!.SlicerCachePivotTables
                    ?? existing.SlicerCacheDefinition.AppendChild(new X14.SlicerCachePivotTables());
                if (!targets.Elements<X14.SlicerCachePivotTable>().Any(item =>
                    string.Equals(item.Name?.Value, pivot.Name, StringComparison.OrdinalIgnoreCase))) {
                    targets.Append(new X14.SlicerCachePivotTable { TabId = ResolvePivotInteractionSheetId(pivot), Name = pivot.Name });
                    existing.SlicerCacheDefinition.Save();
                }
                EnsurePivotCacheInteractionExtensions(pivotPart, pivot.CacheId, slicer: true, timeline: false);
                return existing;
            }
            if (!string.IsNullOrWhiteSpace(requestedCacheName)) {
                throw new InvalidOperationException($"Native slicer cache '{requestedCacheName}' was not found with the requested PivotTable binding.");
            }

            string name = ResolveNativeInteractionName(null, "SlicerCache_" + sourceField, WorkbookPartRoot.SlicerCacheParts
                .Select(part => part.SlicerCacheDefinition?.Name?.Value ?? string.Empty));
            CacheField cacheField = GetPivotSourceCacheField(pivotPart, sourceField, out _);
            var items = new X14.TabularSlicerCacheItems();
            int itemCount = cacheField.SharedItems?.ChildElements.Count ?? 0;
            for (int index = 0; index < itemCount; index++) {
                items.Append(new X14.TabularSlicerCacheItem { Atom = (uint)index, IsSelected = true });
            }
            items.Count = (uint)itemCount;

            SlicerCachePart part = WorkbookPartRoot.AddNewPart<SlicerCachePart>();
            part.SlicerCacheDefinition = new X14.SlicerCacheDefinition(
                new X14.SlicerCachePivotTables(
                    new X14.SlicerCachePivotTable { TabId = ResolvePivotInteractionSheetId(pivot), Name = pivot.Name }),
                new X14.SlicerCacheData(
                    new X14.TabularSlicerCache(items) {
                        PivotCacheId = pivot.CacheId,
                        SortOrder = X14.TabularSlicerCacheSortOrderValues.Ascending,
                        CrossFilter = X14.SlicerCacheCrossFilterValues.ShowItemsWithDataAtTop
                    })) {
                Name = name,
                SourceName = sourceField
            };
            part.SlicerCacheDefinition.Save();
            EnsureWorkbookSlicerReference(part);
            EnsurePivotCacheInteractionExtensions(pivotPart, pivot.CacheId, slicer: true, timeline: false);
            return part;
        }

        private TimeLineCachePart ResolveOrCreateTimelineCache(
            ExcelPivotTableInfo pivot,
            PivotTablePart pivotPart,
            string sourceField,
            string? requestedCacheName,
            TimeLineCachePart? existing,
            (DateTime Minimum, DateTime Maximum)? sourceBounds) {
            if (existing != null) {
                X15.TimelineCachePivotTables targets = existing.TimelineCacheDefinition!.TimelineCachePivotTables
                    ?? existing.TimelineCacheDefinition.AppendChild(new X15.TimelineCachePivotTables());
                if (!targets.Elements<X15.TimelineCachePivotTable>().Any(item =>
                    string.Equals(item.Name?.Value, pivot.Name, StringComparison.OrdinalIgnoreCase))) {
                    targets.Append(new X15.TimelineCachePivotTable { TabId = ResolvePivotInteractionSheetId(pivot), Name = pivot.Name });
                    existing.TimelineCacheDefinition.Save();
                }
                EnsurePivotCacheInteractionExtensions(pivotPart, pivot.CacheId, slicer: false, timeline: true);
                return existing;
            }
            if (!string.IsNullOrWhiteSpace(requestedCacheName)) {
                throw new InvalidOperationException($"Native timeline cache '{requestedCacheName}' was not found with the requested PivotTable binding.");
            }

            string name = ResolveNativeInteractionName(null, "TimelineCache_" + sourceField, WorkbookPartRoot.TimeLineCacheParts
                .Select(part => part.TimelineCacheDefinition?.Name?.Value ?? string.Empty));
            if (!sourceBounds.HasValue) {
                throw new InvalidOperationException("Timeline source bounds were not resolved.");
            }
            (DateTime minimum, DateTime maximum) = sourceBounds.Value;
            TimeLineCachePart part = WorkbookPartRoot.AddNewPart<TimeLineCachePart>();
            part.TimelineCacheDefinition = new X15.TimelineCacheDefinition(
                new X15.TimelineCachePivotTables(
                    new X15.TimelineCachePivotTable { TabId = ResolvePivotInteractionSheetId(pivot), Name = pivot.Name }),
                new X15.TimelineState(
                    new X15.SelectionTimelineRange { StartDate = minimum, EndDate = maximum },
                    new X15.BoundsTimelineRange { StartDate = minimum, EndDate = maximum }) {
                    SingleRangeFilterState = true,
                    MinimalRefreshVersion = 0U,
                    LastRefreshVersion = 0U,
                    PivotCacheId = pivot.CacheId,
                    FilterType = PivotFilterValues.DateBetween
                }) {
                Name = name,
                SourceName = sourceField
            };
            part.TimelineCacheDefinition.Save();
            EnsureWorkbookTimelineReference(part);
            EnsurePivotCacheInteractionExtensions(pivotPart, pivot.CacheId, slicer: false, timeline: true);
            return part;
        }

        private SlicerCachePart? FindCompatibleSlicerCache(ExcelPivotTableInfo pivot, string sourceField, string? requestedName) {
            return WorkbookPartRoot.SlicerCacheParts.FirstOrDefault(part => {
                X14.SlicerCacheDefinition? root = part.SlicerCacheDefinition;
                if (root == null || !string.Equals(root.SourceName?.Value, sourceField, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.IsNullOrWhiteSpace(requestedName)
                    && !string.Equals(root.Name?.Value, requestedName!.Trim(), StringComparison.OrdinalIgnoreCase)) return false;
                return root.SlicerCachePivotTables?.Elements<X14.SlicerCachePivotTable>().Any(item =>
                    PivotInteractionTargetUsesCache(item.Name?.Value, pivot.CacheId)) == true;
            });
        }

        private TimeLineCachePart? FindCompatibleTimelineCache(ExcelPivotTableInfo pivot, string sourceField, string? requestedName) {
            return WorkbookPartRoot.TimeLineCacheParts.FirstOrDefault(part => {
                X15.TimelineCacheDefinition? root = part.TimelineCacheDefinition;
                if (root == null || !string.Equals(root.SourceName?.Value, sourceField, StringComparison.OrdinalIgnoreCase)) return false;
                if (!string.IsNullOrWhiteSpace(requestedName)
                    && !string.Equals(root.Name?.Value, requestedName!.Trim(), StringComparison.OrdinalIgnoreCase)) return false;
                return root.TimelineCachePivotTables?.Elements<X15.TimelineCachePivotTable>().Any(item =>
                    PivotInteractionTargetUsesCache(item.Name?.Value, pivot.CacheId)) == true;
            });
        }

        private bool PivotInteractionTargetUsesCache(string? pivotName, uint cacheId) {
            if (string.IsNullOrWhiteSpace(pivotName)) return false;
            return WorkbookPartRoot.WorksheetParts
                .SelectMany(part => part.PivotTableParts)
                .Any(part => part.PivotTableDefinition?.CacheId?.Value == cacheId
                    && string.Equals(part.PivotTableDefinition?.Name?.Value, pivotName, StringComparison.OrdinalIgnoreCase));
        }

        private uint ResolvePivotInteractionSheetId(ExcelPivotTableInfo pivot) {
            Sheet? sheet = WorkbookPartRoot.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault(item =>
                string.Equals(item.Name?.Value, pivot.SheetName, StringComparison.OrdinalIgnoreCase));
            return sheet?.SheetId?.Value
                ?? throw new InvalidOperationException($"Pivot table worksheet '{pivot.SheetName}' has no stable sheet id.");
        }

        private static CacheField GetPivotSourceCacheField(PivotTablePart pivotPart, string sourceField, out int index) {
            List<CacheField> fields = pivotPart.PivotTableCacheDefinitionPart?.PivotCacheDefinition?
                .CacheFields?.Elements<CacheField>().ToList() ?? new List<CacheField>();
            index = fields.FindIndex(field => string.Equals(field.Name?.Value, sourceField, StringComparison.OrdinalIgnoreCase));
            if (index < 0) throw new InvalidOperationException($"Pivot cache field '{sourceField}' was not found.");
            return fields[index];
        }

        private (DateTime Minimum, DateTime Maximum) GetTimelineBounds(CacheField cacheField, ExcelPivotTableInfo pivot) {
            List<DateTime> values = cacheField.SharedItems?.Elements<DateTimeItem>()
                .Select(item => item.Val?.Value)
                .Where(value => value.HasValue)
                .Select(value => value!.Value)
                .ToList() ?? new List<DateTime>();
            if (values.Count == 0 && !string.IsNullOrWhiteSpace(pivot.SourceSheet)
                && !string.IsNullOrWhiteSpace(pivot.SourceRange)
                && A1.TryParseRange(pivot.SourceRange!.Replace("$", string.Empty), out int r1, out int c1, out int r2, out int c2)) {
                ExcelSheet? sheet = GetPivotInteractionSheetsForCurrentLock().FirstOrDefault(item =>
                    string.Equals(item.Name, pivot.SourceSheet, StringComparison.OrdinalIgnoreCase));
                List<CacheField> fields = FindPivotTablePart(pivot)?.PivotTableCacheDefinitionPart?.PivotCacheDefinition?
                    .CacheFields?.Elements<CacheField>().Where(field => field.DatabaseField?.Value != false).ToList() ?? new List<CacheField>();
                int fieldIndex = fields.FindIndex(field => string.Equals(field.Name?.Value, cacheField.Name?.Value, StringComparison.OrdinalIgnoreCase));
                if (sheet != null && fieldIndex >= 0 && c1 + fieldIndex <= c2) {
                    for (int row = r1 + 1; row <= r2; row++) {
                        if (sheet.TryGetCellValueSnapshot(row, c1 + fieldIndex, out ExcelCellValueSnapshot? snapshot)
                            && snapshot!.DateTimeValue.HasValue) values.Add(snapshot.DateTimeValue.Value);
                    }
                }
            }
            if (values.Count == 0) throw new InvalidOperationException("Timeline source contains no date values.");
            return (values.Min().Date, values.Max().Date);
        }

        private void EnsureWorkbookSlicerReference(SlicerCachePart part) {
            Workbook workbook = WorkbookPartRoot.Workbook
                ?? throw new InvalidDataException("Workbook root is missing.");
            WorkbookExtensionList list = workbook.GetFirstChild<WorkbookExtensionList>()
                ?? workbook.AppendChild(new WorkbookExtensionList());
            WorkbookExtension extension = list.Elements<WorkbookExtension>().FirstOrDefault(item => item.Uri?.Value == SlicerCachesExtensionUri)
                ?? list.AppendChild(new WorkbookExtension { Uri = SlicerCachesExtensionUri });
            X14.SlicerCaches caches = extension.GetFirstChild<X14.SlicerCaches>() ?? extension.AppendChild(new X14.SlicerCaches());
            string id = WorkbookPartRoot.GetIdOfPart(part);
            if (!caches.Elements<X14.SlicerCache>().Any(item => item.Id?.Value == id)) caches.Append(new X14.SlicerCache { Id = id });
            workbook.Save();
        }

        private void EnsureWorkbookTimelineReference(TimeLineCachePart part) {
            Workbook workbook = WorkbookPartRoot.Workbook
                ?? throw new InvalidDataException("Workbook root is missing.");
            WorkbookExtensionList list = workbook.GetFirstChild<WorkbookExtensionList>()
                ?? workbook.AppendChild(new WorkbookExtensionList());
            WorkbookExtension extension = list.Elements<WorkbookExtension>().FirstOrDefault(item => item.Uri?.Value == TimelineCachesExtensionUri)
                ?? list.AppendChild(new WorkbookExtension { Uri = TimelineCachesExtensionUri });
            X15.TimelineCacheReferences caches = extension.GetFirstChild<X15.TimelineCacheReferences>()
                ?? extension.AppendChild(new X15.TimelineCacheReferences());
            string id = WorkbookPartRoot.GetIdOfPart(part);
            if (!caches.Elements<X15.TimelineCacheReference>().Any(item => item.Id?.Value == id)) {
                caches.Append(new X15.TimelineCacheReference { Id = id });
            }
            workbook.Save();
        }

        private static void EnsureWorksheetSlicerReference(WorksheetPart worksheetPart, SlicersPart part) {
            Worksheet worksheet = worksheetPart.Worksheet
                ?? throw new InvalidDataException("Worksheet root is missing.");
            WorksheetExtensionList list = worksheet.GetFirstChild<WorksheetExtensionList>()
                ?? worksheet.AppendChild(new WorksheetExtensionList());
            WorksheetExtension extension = list.Elements<WorksheetExtension>().FirstOrDefault(item => item.Uri?.Value == SlicerListExtensionUri)
                ?? list.AppendChild(new WorksheetExtension { Uri = SlicerListExtensionUri });
            X14.SlicerList refs = extension.GetFirstChild<X14.SlicerList>() ?? extension.AppendChild(new X14.SlicerList());
            string id = worksheetPart.GetIdOfPart(part);
            if (!refs.Elements<X14.SlicerRef>().Any(item => item.Id?.Value == id)) refs.Append(new X14.SlicerRef { Id = id });
        }

        private static void EnsureWorksheetTimelineReference(WorksheetPart worksheetPart, TimeLinePart part) {
            Worksheet worksheet = worksheetPart.Worksheet
                ?? throw new InvalidDataException("Worksheet root is missing.");
            WorksheetExtensionList list = worksheet.GetFirstChild<WorksheetExtensionList>()
                ?? worksheet.AppendChild(new WorksheetExtensionList());
            WorksheetExtension extension = list.Elements<WorksheetExtension>().FirstOrDefault(item => item.Uri?.Value == TimelineListExtensionUri)
                ?? list.AppendChild(new WorksheetExtension { Uri = TimelineListExtensionUri });
            X15.TimelineReferences refs = extension.GetFirstChild<X15.TimelineReferences>()
                ?? extension.AppendChild(new X15.TimelineReferences());
            string id = worksheetPart.GetIdOfPart(part);
            if (!refs.Elements<X15.TimelineReference>().Any(item => item.Id?.Value == id)) refs.Append(new X15.TimelineReference { Id = id });
        }

        private static void EnsurePivotCacheInteractionExtensions(
            PivotTablePart pivotPart,
            uint pivotCacheId,
            bool slicer,
            bool timeline) {
            PivotCacheDefinition definition = pivotPart.PivotTableCacheDefinitionPart?.PivotCacheDefinition
                ?? throw new InvalidOperationException("Pivot cache definition is missing.");
            PivotCacheDefinitionExtensionList list = definition.PivotCacheDefinitionExtensionList
                ?? definition.AppendChild(new PivotCacheDefinitionExtensionList());
            if (slicer) {
                PivotCacheDefinitionExtension extension = list.Elements<PivotCacheDefinitionExtension>()
                    .FirstOrDefault(item => item.Uri?.Value == PivotSlicerExtensionUri)
                    ?? list.AppendChild(new PivotCacheDefinitionExtension { Uri = PivotSlicerExtensionUri });
                X14.PivotCacheDefinition marker = extension.GetFirstChild<X14.PivotCacheDefinition>()
                    ?? extension.AppendChild(new X14.PivotCacheDefinition());
                marker.PivotCacheId = pivotCacheId;
                marker.SlicerData = true;
            }
            if (timeline) {
                PivotCacheDefinitionExtension extension = list.Elements<PivotCacheDefinitionExtension>()
                    .FirstOrDefault(item => item.Uri?.Value == PivotTimelineExtensionUri)
                    ?? list.AppendChild(new PivotCacheDefinitionExtension { Uri = PivotTimelineExtensionUri });
                X15.TimelinePivotCacheDefinition marker = extension.GetFirstChild<X15.TimelinePivotCacheDefinition>()
                    ?? extension.AppendChild(new X15.TimelinePivotCacheDefinition());
                marker.TimelineData = true;
            }
            definition.Save();
        }

    }
}
