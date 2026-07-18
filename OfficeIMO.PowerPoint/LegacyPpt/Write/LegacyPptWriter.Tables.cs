using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Text;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort TablePropertiesIdForWrite = 0x039F;
        private const ushort TableRowPropertiesIdForWrite = 0x03A0;

        internal static bool TryReadTableForWrite(PowerPointTable table,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets,
            out string? reason) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (fonts == null) throw new ArgumentNullException(nameof(fonts));
            if (pictureBullets == null) throw new ArgumentNullException(
                nameof(pictureBullets));
            if (table.Rows <= 0 || table.Columns <= 0
                || table.Width <= 0L || table.Height <= 0L) {
                reason = "A native binary table requires positive bounds, rows, and columns.";
                return false;
            }
            if (!TryReadShapeFormatting(table, out _, out reason)) return false;
            foreach (PowerPointTableColumn column in table.ColumnItems) {
                if (!column.WidthEmus.HasValue || column.WidthEmus.Value <= 0L
                    || ToMasterUnits(column.WidthEmus.Value) <= 0) {
                    reason = "Every native binary table column requires a positive representable width.";
                    return false;
                }
            }
            foreach (PowerPointTableRow row in table.RowItems) {
                if (!row.HeightEmus.HasValue || row.HeightEmus.Value <= 0L
                    || ToMasterUnits(row.HeightEmus.Value) <= 0) {
                    reason = "Every native binary table row requires a positive representable height.";
                    return false;
                }
            }
            for (int row = 0; row < table.Rows; row++) {
                for (int column = 0; column < table.Columns; column++) {
                    PowerPointTableCell cell = table.GetCell(row, column);
                    if (cell.IsMergedCell || cell.Merge is not (1, 1)) {
                        reason = "Merged DrawingML table cells do not yet have a lossless native binary table mapping.";
                        return false;
                    }
                    if (!TryBuildTableCellContent(cell, fonts,
                            pictureBullets, out _, out reason)) {
                        return false;
                    }
                    if (cell.FillColor != null
                        && !OfficeColor.TryParse(cell.FillColor, out _)) {
                        reason = "A table cell fill must be an explicit RGB color for binary PowerPoint.";
                        return false;
                    }
                    if (cell.BorderColor != null
                        && !OfficeColor.TryParse(cell.BorderColor, out _)) {
                        reason = "A table cell border must be an explicit RGB color for binary PowerPoint.";
                        return false;
                    }
                    OfficeBorderBox borders = PowerPointSlideImageRenderer
                        .ResolveTableCellBordersForBinary(table, row,
                            column);
                    if (!CanWriteTableBorder(borders.Left)
                        || !CanWriteTableBorder(borders.Top)
                        || !CanWriteTableBorder(borders.Right)
                        || !CanWriteTableBorder(borders.Bottom)
                        || borders.DiagonalDown?.IsVisible == true
                        || borders.DiagonalUp?.IsVisible == true) {
                        reason = "Table borders require solid single lines without diagonal edges for native binary writing.";
                        return false;
                    }
                }
            }
            for (int rowBoundary = 1; rowBoundary < table.Rows;
                 rowBoundary++) {
                for (int column = 0; column < table.Columns; column++) {
                    OfficeBorderSide? upper = PowerPointSlideImageRenderer
                        .ResolveTableCellBordersForBinary(table,
                            rowBoundary - 1, column).Bottom;
                    OfficeBorderSide? lower = PowerPointSlideImageRenderer
                        .ResolveTableCellBordersForBinary(table,
                            rowBoundary, column).Top;
                    if (!CanShareTableGridEdge(upper, lower)) {
                        reason = "Opposing table cell borders on a shared horizontal edge must use identical color, width, and visibility for native binary writing.";
                        return false;
                    }
                }
            }
            for (int columnBoundary = 1;
                 columnBoundary < table.Columns; columnBoundary++) {
                for (int row = 0; row < table.Rows; row++) {
                    OfficeBorderSide? left = PowerPointSlideImageRenderer
                        .ResolveTableCellBordersForBinary(table, row,
                            columnBoundary - 1).Right;
                    OfficeBorderSide? right = PowerPointSlideImageRenderer
                        .ResolveTableCellBordersForBinary(table, row,
                            columnBoundary).Left;
                    if (!CanShareTableGridEdge(left, right)) {
                        reason = "Opposing table cell borders on a shared vertical edge must use identical color, width, and visibility for native binary writing.";
                        return false;
                    }
                }
            }
            reason = null;
            return true;
        }

        private static int CountTableDrawingShapes(PowerPointTable table) =>
            checked(1 + table.Rows * table.Columns
                + (table.Rows + 1) * table.Columns
                + (table.Columns + 1) * table.Rows);

        private static byte[] BuildTableRecord(PowerPointTable table,
            ref uint nextShapeId,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterAnimationCatalog animationCatalog,
            LegacyPptWriterShapeContext shapeContext,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog? pictureBullets) {
            LegacyPptWriterPictureBulletCatalog bullets = pictureBullets
                ?? LegacyPptWriterPictureBulletCatalog.Empty;
            if (!TryReadTableForWrite(table, fonts, bullets,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            int[] columnWidths = table.ColumnItems.Select(column =>
                ToMasterUnits(column.WidthEmus!.Value)).ToArray();
            int[] rowHeights = table.RowItems.Select(row =>
                ToMasterUnits(row.HeightEmus!.Value)).ToArray();
            int tableWidth = checked(columnWidths.Sum());
            int tableHeight = checked(rowHeights.Sum());

            uint groupShapeId = nextShapeId++;
            var descriptorChildren = new List<byte[]> {
                BuildTableCoordinateRecord(tableWidth, tableHeight),
                BuildTableGroupFsp(table, groupShapeId),
                BuildTablePrimaryPropertiesRecord(table),
                BuildTablePropertiesRecord(rowHeights),
                BuildAnchor(table)
            };
            LegacyPptWriterShapeInteractions interactions =
                interactionCatalog.Get(table);
            byte[]? clientData = BuildClientData(table,
                interactions.ShapeInteractions, animationCatalog.Get(table),
                shapeContext, additionalRecord:
                    BuildTableStyleProgrammableTagsRecord(table));
            if (clientData != null) descriptorChildren.Add(clientData);

            var children = new List<byte[]>(CountTableDrawingShapes(table)) {
                BuildContainer(OfficeArtSpContainer, instance: 0,
                    descriptorChildren)
            };
            int top = 0;
            for (int row = 0; row < table.Rows; row++) {
                int left = 0;
                for (int column = 0; column < table.Columns; column++) {
                    PowerPointTableCell cell = table.GetCell(row, column);
                    children.Add(BuildTableCellRecord(cell,
                        nextShapeId++, row, column, table, left, top,
                        columnWidths[column],
                        rowHeights[row], interactionCatalog, fonts, bullets));
                    left = checked(left + columnWidths[column]);
                }
                top = checked(top + rowHeights[row]);
            }
            AddTableGridLineRecords(children, ref nextShapeId, table,
                columnWidths, rowHeights);
            return BuildContainer(OfficeArtSpgrContainer, instance: 0,
                children);
        }

        private static byte[] BuildTableCoordinateRecord(int width,
            int height) {
            var payload = new byte[16];
            WriteInt32(payload, 8, width);
            WriteInt32(payload, 12, height);
            return BuildRecord(version: 1, instance: 0,
                OfficeArtFspgr, payload);
        }

        private static byte[] BuildTablePropertiesRecord(
            IReadOnlyList<int> rowHeights) {
            var properties = new List<LegacyPptWriterFoptProperty>();
            properties.Add(new LegacyPptWriterFoptProperty(
                TablePropertiesIdForWrite, 1U));
            var rows = new byte[checked(6 + rowHeights.Count * 4)];
            WriteUInt16(rows, 0, checked((ushort)rowHeights.Count));
            WriteUInt16(rows, 2, checked((ushort)rowHeights.Count));
            WriteUInt16(rows, 4, 4);
            for (int index = 0; index < rowHeights.Count; index++) {
                WriteInt32(rows, 6 + index * 4, rowHeights[index]);
            }
            properties.Add(new LegacyPptWriterFoptProperty(
                unchecked((ushort)(0x8000 | TableRowPropertiesIdForWrite)),
                checked((uint)rows.Length), rows));
            return BuildFoptRecord(properties, OfficeArtTertiaryFopt);
        }

        private static byte[] BuildTableStyleProgrammableTagsRecord(
            PowerPointTable table) {
            byte flags = 0;
            if (table.FirstRow) flags |= 1 << 0;
            if (table.LastRow) flags |= 1 << 1;
            if (table.FirstColumn) flags |= 1 << 2;
            if (table.LastColumn) flags |= 1 << 3;
            if (table.BandedRows) flags |= 1 << 4;
            if (table.BandedColumns) flags |= 1 << 5;
            byte[] tagName = BuildRecord(version: 0, instance: 0,
                RecordCString,
                Encoding.Unicode.GetBytes(
                    LegacyPptTableMetadata.TagName));
            byte[] data = BuildRecord(version: 0, instance: 0,
                RecordBinaryTagDataBlob, new byte[] {
                    LegacyPptTableMetadata.Version, flags
                });
            byte[] tag = BuildContainer(RecordProgBinaryTag, instance: 0,
                new[] { tagName, data });
            return BuildContainer(RecordProgTags, instance: 0,
                new[] { tag });
        }

        private static byte[] BuildTablePrimaryPropertiesRecord(
            PowerPointTable table) {
            if (!TryReadShapeFormatting(table,
                    out LegacyPptWriterShapeFormatting formatting,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            var properties = formatting.Properties.ToList();
            properties.Add(new LegacyPptWriterFoptProperty(0x007F,
                0x01000100U));
            return BuildFoptRecord(properties);
        }

        private static byte[] BuildTableGroupFsp(PowerPointTable table,
            uint shapeId) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            uint flags = 0x00000201U;
            if (table.Element.Parent is DocumentFormat.OpenXml.Presentation.GroupShape) {
                flags |= 1U << 1;
            }
            WriteUInt32(payload, 4, flags);
            return BuildRecord(version: 2, instance: 0,
                OfficeArtFsp, payload);
        }

        private static byte[] BuildTableCellRecord(PowerPointTableCell cell,
            uint shapeId, int row, int column, PowerPointTable table,
            int left, int top, int width, int height,
            LegacyPptWriterInteractionCatalog interactionCatalog,
            LegacyPptWriterFontCatalog fonts,
            LegacyPptWriterPictureBulletCatalog pictureBullets) {
            if (!TryBuildTableCellContent(cell, fonts, pictureBullets,
                    out LegacyPptWriterTextBoxContent? textContent,
                    out string? reason)) {
                throw new NotSupportedException(reason);
            }
            var children = new List<byte[]> {
                BuildTableCellFsp(shapeId),
                BuildTableCellFoptRecord(table, row, column, shapeId),
                BuildTableCellAnchor(left, top, width, height)
            };
            if (textContent!.Style9Record != null) {
                children.Add(BuildContainer(OfficeArtClientData, instance: 0,
                    new[] { BuildShapePpt9ProgrammableTagsRecord(
                        textContent.Style9Record) }));
            }
            LegacyPptWriterShapeInteractions interactions =
                interactionCatalog.Get(cell.Cell);
            children.Add(BuildTextBox((uint)LegacyPptTextType.Other,
                interactions.TextInteractions, textContent));
            return BuildContainer(OfficeArtSpContainer, instance: 0,
                children);
        }

        private static byte[] BuildTableCellFsp(uint shapeId) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            WriteUInt32(payload, 4, 0x00000A02U);
            return BuildRecord(version: 2, instance: 1,
                OfficeArtFsp, payload);
        }

        private static byte[] BuildTableCellFoptRecord(
            PowerPointTable table, int row, int column, uint shapeId) {
            OfficeColor fill = PowerPointSlideImageRenderer
                .ResolveTableCellFillColorForBinary(table, row, column);
            var properties = new List<LegacyPptWriterFoptProperty> {
                new(0x0080, shapeId),
                new(0x0081, 91440U),
                new(0x0082, 45720U),
                new(0x0083, 91440U),
                new(0x0084, 45720U),
                new(0x0085, 0U),
                new(0x0087, 0U),
                new(0x00BF, 0x00040004U),
                new(0x0180, 0U),
                new(0x0181, PackOfficeArtColor(fill)),
                new(0x0183, PackOfficeArtColor(fill)),
                new(0x01BF, 0x00100010U),
                new(0x01D6, 1U),
                new(0x01FF, 0x00090000U)
            };
            return BuildFoptRecord(properties);
        }

        private static void AddTableGridLineRecords(
            ICollection<byte[]> children, ref uint nextShapeId,
            PowerPointTable table,
            IReadOnlyList<int> columnWidths,
            IReadOnlyList<int> rowHeights) {
            int top = 0;
            for (int rowBoundary = 0;
                 rowBoundary <= rowHeights.Count; rowBoundary++) {
                int left = 0;
                for (int column = 0; column < columnWidths.Count; column++) {
                    children.Add(BuildTableGridLineRecord(nextShapeId++,
                        left, top, columnWidths[column], 0,
                        GetHorizontalTableBorder(table, rowBoundary,
                            column)));
                    left = checked(left + columnWidths[column]);
                }
                if (rowBoundary < rowHeights.Count) {
                    top = checked(top + rowHeights[rowBoundary]);
                }
            }
            int lineLeft = 0;
            for (int columnBoundary = 0;
                 columnBoundary <= columnWidths.Count; columnBoundary++) {
                int lineTop = 0;
                for (int row = 0; row < rowHeights.Count; row++) {
                    children.Add(BuildTableGridLineRecord(nextShapeId++,
                        lineLeft, lineTop, 0, rowHeights[row],
                        GetVerticalTableBorder(table, row,
                            columnBoundary)));
                    lineTop = checked(lineTop + rowHeights[row]);
                }
                if (columnBoundary < columnWidths.Count) {
                    lineLeft = checked(lineLeft
                        + columnWidths[columnBoundary]);
                }
            }
        }

        private static byte[] BuildTableGridLineRecord(uint shapeId,
            int left, int top, int width, int height,
            OfficeBorderSide? border) {
            OfficeBorderSide side = border
                ?? new OfficeBorderSide(OfficeColor.Black, 0D);
            var properties = new List<LegacyPptWriterFoptProperty> {
                new(0x0144, 4U),
                new(0x01C0, PackOfficeArtColor(side.Color)),
                new(0x01CB, checked((uint)Math.Round(
                    side.Width * 12700D,
                    MidpointRounding.AwayFromZero))),
                new(0x01FF, side.IsVisible
                    ? 0x000A0008U
                    : 0x00080000U),
                new(0x023F, 0x00020000U),
                new(0x02BF, 0x00080000U)
            };
            return BuildContainer(OfficeArtSpContainer, instance: 0,
                new[] {
                    BuildTableLineFsp(shapeId),
                    BuildFoptRecord(properties),
                    BuildTableCellAnchor(left, top, width, height)
                });
        }

        private static OfficeBorderSide? GetHorizontalTableBorder(
            PowerPointTable table, int rowBoundary, int column) {
            if (rowBoundary == 0) {
                return PowerPointSlideImageRenderer
                    .ResolveTableCellBordersForBinary(table, 0, column).Top;
            }
            return PowerPointSlideImageRenderer
                .ResolveTableCellBordersForBinary(table,
                    rowBoundary - 1, column).Bottom;
        }

        private static OfficeBorderSide? GetVerticalTableBorder(
            PowerPointTable table, int row, int columnBoundary) {
            if (columnBoundary == 0) {
                return PowerPointSlideImageRenderer
                    .ResolveTableCellBordersForBinary(table, row, 0).Left;
            }
            return PowerPointSlideImageRenderer
                .ResolveTableCellBordersForBinary(table, row,
                    columnBoundary - 1).Right;
        }

        private static bool CanWriteTableBorder(OfficeBorderSide? border) =>
            !border.HasValue || !border.Value.IsVisible
            || border.Value.DashStyle == OfficeStrokeDashStyle.Solid
            && border.Value.LineKind == OfficeBorderLineKind.Single;

        private static bool CanShareTableGridEdge(
            OfficeBorderSide? first, OfficeBorderSide? second) {
            bool firstVisible = first?.IsVisible == true;
            bool secondVisible = second?.IsVisible == true;
            if (!firstVisible && !secondVisible) return true;
            return firstVisible && secondVisible
                && first!.Value == second!.Value;
        }

        private static byte[] BuildTableLineFsp(uint shapeId) {
            var payload = new byte[8];
            WriteUInt32(payload, 0, shapeId);
            WriteUInt32(payload, 4, 0x00000A02U);
            return BuildRecord(version: 2, instance: 20,
                OfficeArtFsp, payload);
        }

        private static byte[] BuildTableCellAnchor(int left, int top,
            int width, int height) {
            var payload = new byte[16];
            WriteInt32(payload, 0, left);
            WriteInt32(payload, 4, top);
            WriteInt32(payload, 8, checked(left + width));
            WriteInt32(payload, 12, checked(top + height));
            return BuildRecord(version: 0, instance: 0,
                OfficeArtChildAnchor, payload);
        }
    }
}
