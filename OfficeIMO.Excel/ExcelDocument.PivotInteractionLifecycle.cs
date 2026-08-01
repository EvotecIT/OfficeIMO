using System.Globalization;
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
        private void PreparePivotInteractionsForWorksheetRemoval(string worksheetName) {
            ExcelSheet? removedSheet = Sheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase));
            if (removedSheet == null) return;
            var removedPivotNames = new HashSet<string>(
                removedSheet.GetPivotTables().Select(pivot => pivot.Name),
                StringComparer.OrdinalIgnoreCase);
            ExcelPivotInteractionInfo[] interactions = GetPivotInteractions().ToArray();
            var interactionNamesToRemove = new HashSet<string>(
                interactions.Where(item => string.Equals(item.WorksheetName, removedSheet.Name, StringComparison.OrdinalIgnoreCase))
                    .Select(item => item.Name),
                StringComparer.OrdinalIgnoreCase);
            var emptyCaches = new List<(ExcelPivotInteractionCacheKind Kind, string Name)>();

            foreach (SlicerCachePart part in WorkbookPartRoot.SlicerCacheParts.ToList()) {
                X14.SlicerCachePivotTables? targets = part.SlicerCacheDefinition?.SlicerCachePivotTables;
                X14.SlicerCachePivotTable[] removedTargets = targets?.Elements<X14.SlicerCachePivotTable>()
                    .Where(item => removedPivotNames.Contains(item.Name?.Value ?? string.Empty)).ToArray()
                    ?? Array.Empty<X14.SlicerCachePivotTable>();
                if (removedTargets.Length == 0) continue;
                string cacheName = part.SlicerCacheDefinition?.Name?.Value ?? string.Empty;
                if (targets!.Elements<X14.SlicerCachePivotTable>().Count() == removedTargets.Length) {
                    emptyCaches.Add((ExcelPivotInteractionCacheKind.Slicer, cacheName));
                    foreach (ExcelPivotInteractionInfo interaction in interactions.Where(item =>
                        item.Kind == ExcelPivotInteractionCacheKind.Slicer
                        && string.Equals(item.CacheName, cacheName, StringComparison.OrdinalIgnoreCase))) {
                        interactionNamesToRemove.Add(interaction.Name);
                    }
                } else {
                    foreach (X14.SlicerCachePivotTable target in removedTargets) target.Remove();
                    part.SlicerCacheDefinition!.Save();
                }
            }

            foreach (TimeLineCachePart part in WorkbookPartRoot.TimeLineCacheParts.ToList()) {
                X15.TimelineCachePivotTables? targets = part.TimelineCacheDefinition?.TimelineCachePivotTables;
                X15.TimelineCachePivotTable[] removedTargets = targets?.Elements<X15.TimelineCachePivotTable>()
                    .Where(item => removedPivotNames.Contains(item.Name?.Value ?? string.Empty)).ToArray()
                    ?? Array.Empty<X15.TimelineCachePivotTable>();
                if (removedTargets.Length == 0) continue;
                string cacheName = part.TimelineCacheDefinition?.Name?.Value ?? string.Empty;
                if (targets!.Elements<X15.TimelineCachePivotTable>().Count() == removedTargets.Length) {
                    emptyCaches.Add((ExcelPivotInteractionCacheKind.Timeline, cacheName));
                    foreach (ExcelPivotInteractionInfo interaction in interactions.Where(item =>
                        item.Kind == ExcelPivotInteractionCacheKind.Timeline
                        && string.Equals(item.CacheName, cacheName, StringComparison.OrdinalIgnoreCase))) {
                        interactionNamesToRemove.Add(interaction.Name);
                    }
                } else {
                    foreach (X15.TimelineCachePivotTable target in removedTargets) target.Remove();
                    part.TimelineCacheDefinition!.Save();
                }
            }

            foreach (string interactionName in interactionNamesToRemove) RemovePivotInteraction(interactionName);
            foreach ((ExcelPivotInteractionCacheKind kind, string cacheName) in emptyCaches) {
                bool exists = kind == ExcelPivotInteractionCacheKind.Slicer
                    ? WorkbookPartRoot.SlicerCacheParts.Any(part =>
                        string.Equals(part.SlicerCacheDefinition?.Name?.Value, cacheName, StringComparison.OrdinalIgnoreCase))
                    : WorkbookPartRoot.TimeLineCacheParts.Any(part =>
                        string.Equals(part.TimelineCacheDefinition?.Name?.Value, cacheName, StringComparison.OrdinalIgnoreCase));
                if (exists) RemoveNativeInteractionCache(kind, cacheName);
            }
        }

        private static void AddPivotInteractionDrawing(
            ExcelSheet sheet,
            string name,
            int row,
            int column,
            int widthPixels,
            int heightPixels,
            bool timeline) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart;
            Worksheet worksheet = worksheetPart.Worksheet
                ?? throw new InvalidDataException("Worksheet root is missing.");
            DocumentFormat.OpenXml.Spreadsheet.Drawing? drawing = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            if (drawing == null) {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            } else {
                drawingsPart = (DrawingsPart)worksheetPart.GetPartById(drawing.Id!);
                drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            }

            long width = (long)Math.Round(widthPixels * 9525D);
            long height = (long)Math.Round(heightPixels * 9525D);
            UInt32Value id = NextPivotInteractionDrawingId(drawingsPart);
            OpenXmlElement graphicReference = timeline
                ? new X15TimelineDrawing.TimeSlicer { Name = name }
                : new X14SlicerDrawing.Slicer { Name = name };
            string graphicUri = timeline ? TimelineGraphicDataUri : SlicerGraphicDataUri;
            var frame = new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = id, Name = name },
                    new Xdr.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks { NoChangeAspect = true })),
                new Xdr.Transform(new Offset { X = 0, Y = 0 }, new Extents { Cx = width, Cy = height }),
                new Graphic(new GraphicData(graphicReference) { Uri = graphicUri }));
            drawingsPart.WorksheetDrawing!.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = width, Cy = height },
                frame,
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing!.Save();
        }

        private static UInt32Value NextPivotInteractionDrawingId(DrawingsPart drawingsPart) {
            uint maximum = drawingsPart.WorksheetDrawing?.Descendants<Xdr.NonVisualDrawingProperties>()
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max() ?? 0U;
            return maximum + 1U;
        }

        private static void RemovePivotInteractionDrawing(ExcelSheet sheet, string name, bool timeline) {
            DrawingsPart? drawingsPart = sheet.WorksheetPart.DrawingsPart;
            if (drawingsPart?.WorksheetDrawing == null) return;
            Xdr.GraphicFrame? frame = drawingsPart.WorksheetDrawing.Descendants<Xdr.GraphicFrame>().FirstOrDefault(candidate => {
                string? candidateName = timeline
                    ? candidate.Graphic?.GraphicData?.GetFirstChild<X15TimelineDrawing.TimeSlicer>()?.Name?.Value
                    : candidate.Graphic?.GraphicData?.GetFirstChild<X14SlicerDrawing.Slicer>()?.Name?.Value;
                return string.Equals(candidateName, name, StringComparison.OrdinalIgnoreCase);
            });
            frame?.Parent?.Remove();
            if (drawingsPart.WorksheetDrawing.ChildElements.Any()) {
                drawingsPart.WorksheetDrawing.Save();
                return;
            }
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException("Worksheet root is missing.");
            worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>()?.Remove();
            sheet.WorksheetPart.DeletePart(drawingsPart);
            worksheet.Save();
        }

        private static void CleanupSlicerViewPart(WorksheetPart worksheetPart, SlicersPart part) {
            if (part.Slicers?.Elements<X14.Slicer>().Any() == true) {
                part.Slicers.Save();
                return;
            }
            string id = worksheetPart.GetIdOfPart(part);
            RemoveWorksheetPartReference<X14.SlicerList, X14.SlicerRef>(
                worksheetPart.Worksheet ?? throw new InvalidDataException("Worksheet root is missing."),
                SlicerListExtensionUri,
                id);
            worksheetPart.DeletePart(part);
        }

        private static void CleanupTimelineViewPart(WorksheetPart worksheetPart, TimeLinePart part) {
            if (part.Timelines?.Elements<X15.Timeline>().Any() == true) {
                part.Timelines.Save();
                return;
            }
            string id = worksheetPart.GetIdOfPart(part);
            RemoveWorksheetPartReference<X15.TimelineReferences, X15.TimelineReference>(
                worksheetPart.Worksheet ?? throw new InvalidDataException("Worksheet root is missing."),
                TimelineListExtensionUri,
                id);
            worksheetPart.DeletePart(part);
        }

        private static void RemoveWorksheetPartReference<TList, TReference>(Worksheet worksheet, string uri, string relationshipId)
            where TList : OpenXmlCompositeElement
            where TReference : OpenXmlElement {
            WorksheetExtensionList? extensionList = worksheet.GetFirstChild<WorksheetExtensionList>();
            WorksheetExtension? extension = extensionList?.Elements<WorksheetExtension>()
                .FirstOrDefault(item => item.Uri?.Value == uri);
            TList? list = extension?.GetFirstChild<TList>();
            TReference? reference = list?.Elements<TReference>().FirstOrDefault(item =>
                item.GetAttribute("id", OfficeDocumentRelationshipsNamespace).Value == relationshipId);
            reference?.Remove();
            if (list != null && !list.ChildElements.Any()) list.Remove();
            if (extension != null && !extension.ChildElements.Any()) extension.Remove();
            if (extensionList != null && !extensionList.ChildElements.Any()) extensionList.Remove();
        }

        private void RemoveNativeInteractionCache(ExcelPivotInteractionCacheKind kind, string cacheName) {
            string[] pivotNames;
            if (kind == ExcelPivotInteractionCacheKind.Slicer) {
                SlicerCachePart? part = WorkbookPartRoot.SlicerCacheParts.FirstOrDefault(candidate =>
                    string.Equals(candidate.SlicerCacheDefinition?.Name?.Value, cacheName, StringComparison.OrdinalIgnoreCase));
                if (part == null) return;
                pivotNames = part.SlicerCacheDefinition?.SlicerCachePivotTables?
                    .Elements<X14.SlicerCachePivotTable>()
                    .Select(item => item.Name?.Value ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray() ?? Array.Empty<string>();
                string id = WorkbookPartRoot.GetIdOfPart(part);
                RemoveWorkbookPartReference<X14.SlicerCaches, X14.SlicerCache>(SlicerCachesExtensionUri, id);
                WorkbookPartRoot.DeletePart(part);
            } else {
                TimeLineCachePart? part = WorkbookPartRoot.TimeLineCacheParts.FirstOrDefault(candidate =>
                    string.Equals(candidate.TimelineCacheDefinition?.Name?.Value, cacheName, StringComparison.OrdinalIgnoreCase));
                if (part == null) return;
                pivotNames = part.TimelineCacheDefinition?.TimelineCachePivotTables?
                    .Elements<X15.TimelineCachePivotTable>()
                    .Select(item => item.Name?.Value ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray() ?? Array.Empty<string>();
                string id = WorkbookPartRoot.GetIdOfPart(part);
                RemoveWorkbookPartReference<X15.TimelineCacheReferences, X15.TimelineCacheReference>(TimelineCachesExtensionUri, id);
                WorkbookPartRoot.DeletePart(part);
            }
            uint[] pivotCacheIds = GetPivotTables()
                .Where(pivot => pivotNames.Contains(pivot.Name, StringComparer.OrdinalIgnoreCase))
                .Select(pivot => pivot.CacheId)
                .Distinct()
                .ToArray();
            foreach (uint pivotCacheId in pivotCacheIds) CleanupPivotInteractionExtension(pivotCacheId, kind);
            (WorkbookPartRoot.Workbook ?? throw new InvalidDataException("Workbook root is missing.")).Save();
        }

        private void CleanupPivotInteractionExtension(uint pivotCacheId, ExcelPivotInteractionCacheKind kind) {
            bool stillUsed = kind == ExcelPivotInteractionCacheKind.Slicer
                ? WorkbookPartRoot.SlicerCacheParts.Any(part => part.SlicerCacheDefinition?.SlicerCachePivotTables?
                    .Elements<X14.SlicerCachePivotTable>().Any(item =>
                        PivotInteractionTargetUsesCache(item.Name?.Value, pivotCacheId)) == true)
                : WorkbookPartRoot.TimeLineCacheParts.Any(part => part.TimelineCacheDefinition?.TimelineCachePivotTables?
                    .Elements<X15.TimelineCachePivotTable>().Any(item =>
                        PivotInteractionTargetUsesCache(item.Name?.Value, pivotCacheId)) == true);
            if (stillUsed) return;
            ExcelPivotTableInfo? pivot = GetPivotTables().FirstOrDefault(item =>
                item.CacheId == pivotCacheId);
            PivotCacheDefinition? definition = pivot == null
                ? null
                : FindPivotTablePart(pivot)?.PivotTableCacheDefinitionPart?.PivotCacheDefinition;
            PivotCacheDefinitionExtensionList? list = definition?.PivotCacheDefinitionExtensionList;
            string uri = kind == ExcelPivotInteractionCacheKind.Slicer
                ? PivotSlicerExtensionUri
                : PivotTimelineExtensionUri;
            list?.Elements<PivotCacheDefinitionExtension>()
                .FirstOrDefault(item => item.Uri?.Value == uri)?
                .Remove();
            if (list != null && !list.ChildElements.Any()) list.Remove();
            definition?.Save();
        }

        private void RemoveWorkbookPartReference<TList, TReference>(string uri, string relationshipId)
            where TList : OpenXmlCompositeElement
            where TReference : OpenXmlElement {
            Workbook workbook = WorkbookPartRoot.Workbook
                ?? throw new InvalidDataException("Workbook root is missing.");
            WorkbookExtensionList? extensionList = workbook.GetFirstChild<WorkbookExtensionList>();
            WorkbookExtension? extension = extensionList?.Elements<WorkbookExtension>()
                .FirstOrDefault(item => item.Uri?.Value == uri);
            TList? list = extension?.GetFirstChild<TList>();
            TReference? reference = list?.Elements<TReference>().FirstOrDefault(item =>
                item.GetAttribute("id", OfficeDocumentRelationshipsNamespace).Value == relationshipId);
            reference?.Remove();
            if (list != null && !list.ChildElements.Any()) list.Remove();
            if (extension != null && !extension.ChildElements.Any()) extension.Remove();
            if (extensionList != null && !extensionList.ChildElements.Any()) extensionList.Remove();
        }
    }
}
