using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static PowerPointTable ProjectLegacyTable(PowerPointSlide slide,
            LegacyPptShape source,
            IReadOnlyDictionary<uint, SlidePart> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext) {
            P.ShapeTree tree = slide.SlidePart.Slide?.CommonSlideData?.ShapeTree
                ?? throw new InvalidDataException(
                    "The projected slide has no shape tree.");
            uint shapeId = tree.Descendants<P.NonVisualDrawingProperties>()
                .Select(item => item.Id?.Value ?? 0U)
                .DefaultIfEmpty(1U)
                .Max() + 1U;
            P.GraphicFrame frame = CreateLegacyTableFrame(slide.SlidePart,
                source, shapeId, slidePartsByLegacyId, soundContext);
            tree.Append(frame);
            slide.ReserveShapeIdsThrough(shapeId + 1U);
            return new PowerPointTable(frame, slide.SlidePart);
        }

        private static P.GraphicFrame CreateLegacyTableFrame(
            OpenXmlPart ownerPart, LegacyPptShape source, uint shapeId,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) {
            LegacyPptTable tableSource = source.Table
                ?? throw new ArgumentException("The source shape is not a native binary table.", nameof(source));
            long width = Math.Max(1L, ToEmus(source.Bounds.Width));
            long height = Math.Max(1L, ToEmus(source.Bounds.Height));
            long[] columnWidths = ScaleLegacyTableDimensions(
                tableSource.ColumnWidths, width);
            long[] rowHeights = ScaleLegacyTableDimensions(
                tableSource.RowHeights, height);
            var tableElement = new A.Table(
                new A.TableProperties(),
                new A.TableGrid(columnWidths.Select(value =>
                    new A.GridColumn { Width = value })));
            for (int row = 0; row < tableSource.Rows; row++) {
                var tableRow = new A.TableRow { Height = rowHeights[row] };
                for (int column = 0; column < tableSource.Columns; column++) {
                    tableRow.Append(new A.TableCell(
                        PowerPointTableTextDefaults.CreateTextBody(),
                        new A.TableCellProperties()));
                }
                tableElement.Append(tableRow);
            }
            var frame = new P.GraphicFrame(
                new P.NonVisualGraphicFrameProperties(
                    new P.NonVisualDrawingProperties {
                        Id = shapeId,
                        Name = $"Binary Table {shapeId - 1U}"
                    },
                    new P.NonVisualGraphicFrameDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.Transform(
                    new A.Offset {
                        X = ToEmus(source.Bounds.Left),
                        Y = ToEmus(source.Bounds.Top)
                    },
                    new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(tableElement) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/table"
                }));
            var table = new PowerPointTable(frame,
                ownerPart as SlidePart);
            table.FirstRow = false;
            table.LastRow = false;
            table.FirstColumn = false;
            table.LastColumn = false;
            table.BandedRows = false;
            table.BandedColumns = false;
            table.StyleId = null;

            for (int index = 0; index < columnWidths.Length; index++) {
                table.ColumnItems[index].WidthEmus = columnWidths[index];
            }
            for (int index = 0; index < rowHeights.Length; index++) {
                table.RowItems[index].HeightEmus = rowHeights[index];
            }

            foreach (LegacyPptTableCell sourceCell in tableSource.Cells) {
                PowerPointTableCell targetCell = table.GetCell(
                    sourceCell.Row, sourceCell.Column);
                targetCell.Cell.TextBody = CreateLegacyTableTextBody(
                    ownerPart, sourceCell.SourceShape,
                    slidePartsByLegacyId, soundContext);
                if (sourceCell.SourceShape.FillColor != null) {
                    targetCell.FillColor = sourceCell.SourceShape.FillColor;
                }
                if (sourceCell.SourceShape.LineColor != null) {
                    targetCell.BorderColor = sourceCell.SourceShape.LineColor;
                }
            }
            foreach (LegacyPptTableCell sourceCell in tableSource.Cells.Where(
                         cell => cell.RowSpan > 1 || cell.ColumnSpan > 1)) {
                table.MergeCells(sourceCell.Row, sourceCell.Column,
                    sourceCell.Row + sourceCell.RowSpan - 1,
                    sourceCell.Column + sourceCell.ColumnSpan - 1);
            }
            return frame;
        }

        private static A.TextBody CreateLegacyTableTextBody(
            OpenXmlPart ownerPart, LegacyPptShape source,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) {
            P.TextBody projected = LegacyPptTextProjection.CreateTextBody(
                source.TextBody, source.TextFrame,
                interaction => ProjectLegacyInteraction(ownerPart,
                    interaction, slidePartsByLegacyId: slidePartsByLegacyId,
                    soundContext: soundContext),
                pictureBullet => ProjectLegacyPictureBullet(ownerPart,
                    pictureBullet));
            var body = new A.TextBody();
            foreach (OpenXmlElement child in projected.ChildElements) {
                body.Append(child.CloneNode(true));
            }
            return body;
        }

        private static long[] ScaleLegacyTableDimensions(
            IReadOnlyList<int> source, long targetTotal) {
            long sourceTotal = source.Sum(value => (long)value);
            if (source.Count == 0 || sourceTotal <= 0L) {
                throw new InvalidDataException(
                    "A native binary table has an invalid grid definition.");
            }
            var result = new long[source.Count];
            long assigned = 0L;
            for (int index = 0; index < source.Count; index++) {
                long value = index == source.Count - 1
                    ? targetTotal - assigned
                    : Math.Max(1L, checked((long)Math.Round(
                        source[index] / (double)sourceTotal * targetTotal,
                        MidpointRounding.AwayFromZero)));
                result[index] = value;
                assigned = checked(assigned + value);
            }
            int lastIndex = result.Length - 1;
            if (result[lastIndex] <= 0L) {
                result[lastIndex] = 1L;
                result[0] = Math.Max(1L, result[0] - 1L);
            }
            return result;
        }
    }
}
