using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private readonly struct HeaderFooterImagePlacement {
            internal HeaderFooterImagePlacement(bool isHeader, HeaderFooterPosition position) {
                IsHeader = isHeader;
                Position = position;
            }

            internal bool IsHeader { get; }

            internal HeaderFooterPosition Position { get; }
        }

        private static void ProjectWorksheetImages(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> imageStore = workbook.DrawingRecords
                .Where(record => record.BlipStoreEntries.Count > 0)
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.HasImportableImagePayload)
                .ToArray();
            if (imageStore.Count == 0) {
                return;
            }

            foreach (LegacyXlsDrawingRecord drawing in workbook.DrawingRecords.Where(record =>
                record.Kind == LegacyXlsDrawingRecordKind.Drawing
                && string.Equals(record.SheetName, legacySheet.Name, StringComparison.Ordinal)
                && IsPictureDrawing(record))) {
                LegacyXlsDrawingAnchor? anchor = drawing.AnchorEntries.FirstOrDefault();
                if (anchor == null) {
                    continue;
                }

                LegacyXlsDrawingBlipStoreEntry? image = TryResolveImage(drawing, imageStore);
                if (image?.EmbeddedBlipContentType == null || image.EmbeddedBlipPayloadBytes.Length == 0) {
                    continue;
                }

                ProjectWorksheetImage(sheet, drawing, anchor, image);
            }
        }

        private static void ProjectHeaderFooterImages(LegacyXlsWorkbook workbook, LegacyXlsWorksheet legacySheet, ExcelSheet sheet) {
            LegacyXlsPageSetup? pageSetup = legacySheet.PageSetup;
            if (pageSetup == null || !TryGetSingleHeaderFooterImagePlacement(pageSetup, out HeaderFooterImagePlacement placement)) {
                return;
            }

            IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> imageStore = workbook.DrawingRecords
                .Where(record => record.BlipStoreEntries.Count > 0)
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.HasImportableImagePayload)
                .ToArray();
            if (imageStore.Count == 0) {
                return;
            }

            LegacyXlsDrawingBlipStoreEntry? image = null;
            int imageCount = 0;
            foreach (LegacyXlsDrawingRecord drawing in workbook.DrawingRecords.Where(record =>
                record.Kind == LegacyXlsDrawingRecordKind.HeaderFooterPicture
                && string.Equals(record.SheetName, legacySheet.Name, StringComparison.Ordinal)
                && record.HasSupportedHeaderFooterPictureMetadata
                && IsPictureDrawing(record))) {
                LegacyXlsDrawingBlipStoreEntry? candidate = TryResolveImage(drawing, imageStore);
                if (candidate?.EmbeddedBlipContentType == null || candidate.EmbeddedBlipPayloadBytes.Length == 0) {
                    continue;
                }

                image = candidate;
                imageCount++;
                if (imageCount > 1) {
                    return;
                }
            }

            if (image == null || imageCount != 1) {
                return;
            }

            if (placement.IsHeader) {
                sheet.SetHeaderImage(placement.Position, image.EmbeddedBlipPayloadBytes, image.EmbeddedBlipContentType!);
            } else {
                sheet.SetFooterImage(placement.Position, image.EmbeddedBlipPayloadBytes, image.EmbeddedBlipContentType!);
            }
        }

        private static bool TryGetSingleHeaderFooterImagePlacement(LegacyXlsPageSetup pageSetup, out HeaderFooterImagePlacement placement) {
            placement = default;
            int count = 0;
            AddHeaderFooterImagePlacements(pageSetup.HeaderText, isHeader: true, ref placement, ref count);
            AddHeaderFooterImagePlacements(pageSetup.FooterText, isHeader: false, ref placement, ref count);

            return count == 1;
        }

        private static void AddHeaderFooterImagePlacements(
            string? text,
            bool isHeader,
            ref HeaderFooterImagePlacement placement,
            ref int count) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            (string? Left, string? Center, string? Right) sections = SplitHeaderFooterText(text);
            AddHeaderFooterImagePlacement(sections.Left, isHeader, HeaderFooterPosition.Left, ref placement, ref count);
            AddHeaderFooterImagePlacement(sections.Center, isHeader, HeaderFooterPosition.Center, ref placement, ref count);
            AddHeaderFooterImagePlacement(sections.Right, isHeader, HeaderFooterPosition.Right, ref placement, ref count);
        }

        private static void AddHeaderFooterImagePlacement(
            string? sectionText,
            bool isHeader,
            HeaderFooterPosition position,
            ref HeaderFooterImagePlacement placement,
            ref int count) {
            if (string.IsNullOrEmpty(sectionText) || sectionText!.IndexOf("&G", StringComparison.Ordinal) < 0) {
                return;
            }

            placement = new HeaderFooterImagePlacement(isHeader, position);
            count++;
        }

        private static bool IsPictureDrawing(LegacyXlsDrawingRecord drawing) {
            return drawing.ShapeEntries.Any(shape => shape.ShapeType == 0x004B)
                || drawing.ShapeProperties.Any(property => property.IsBlipId && property.PropertyId == 0x0104);
        }

        private static LegacyXlsDrawingBlipStoreEntry? TryResolveImage(
            LegacyXlsDrawingRecord drawing,
            IReadOnlyList<LegacyXlsDrawingBlipStoreEntry> imageStore) {
            LegacyXlsDrawingShapeProperty? blipProperty = drawing.ShapeProperties
                .FirstOrDefault(property => property.IsBlipId && property.PropertyId == 0x0104 && property.Value > 0);
            if (blipProperty != null) {
                int index = checked((int)blipProperty.Value) - 1;
                if ((uint)index < (uint)imageStore.Count) {
                    return imageStore[index];
                }
            }

            return imageStore.Count == 1 ? imageStore[0] : null;
        }

        private static void ProjectWorksheetImage(
            ExcelSheet sheet,
            LegacyXlsDrawingRecord drawing,
            LegacyXlsDrawingAnchor anchor,
            LegacyXlsDrawingBlipStoreEntry image) {
            int startRow = anchor.StartRow + 1;
            int startColumn = anchor.StartColumn + 1;
            int endRow = Math.Max(startRow, anchor.EndRow + 1);
            int endColumn = Math.Max(startColumn, anchor.EndColumn + 1);
            string name = $"Legacy XLS Picture {drawing.RecordOffset}";

            if (endRow > startRow || endColumn > startColumn) {
                string range = BuildA1Range(startRow, startColumn, endRow, endColumn);
                sheet.AddImageToRange(
                    range,
                    image.EmbeddedBlipPayloadBytes,
                    image.EmbeddedBlipContentType!,
                    name: name,
                    altText: name);
            } else {
                sheet.AddImage(
                    startRow,
                    startColumn,
                    image.EmbeddedBlipPayloadBytes,
                    image.EmbeddedBlipContentType!,
                    name: name,
                    altText: name);
            }
        }
    }
}
