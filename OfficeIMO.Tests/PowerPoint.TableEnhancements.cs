using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTableEnhancementsTests {
        [Fact]
        public void SetRowHeightsEvenly_UsesTableHeight() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 3, columns: 2, left: 0, top: 0, width: 4000, height: 9000);

                table.SetRowHeightsEvenly();

                Assert.Equal(3000, table.GetRowHeight(0));
                Assert.Equal(3000, table.GetRowHeight(1));
                Assert.Equal(3000, table.GetRowHeight(2));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SetRowHeightsByRatio_UsesTableHeight() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 2, columns: 1, left: 0, top: 0, width: 4000, height: 9000);

                table.SetRowHeightsByRatio(1, 2);

                Assert.Equal(3000, table.GetRowHeight(0));
                Assert.Equal(6000, table.GetRowHeight(1));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void TableSizeGetters_ReturnUnits() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 1, columns: 1, left: 0, top: 0, width: 4000, height: 2000);

                table.SetColumnWidthPoints(0, 72);
                table.SetRowHeightCm(0, 1.0);

                Assert.Equal(72, table.GetColumnWidthPoints(0), 3);
                Assert.Equal(1.0, table.GetRowHeightCm(0), 3);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SetCellAlignment_AppliesToRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 2, columns: 2, left: 0, top: 0, width: 4000, height: 3000);

                table.SetCellAlignment(A.TextAlignmentTypeValues.Center, A.TextAnchoringTypeValues.Center,
                    startRow: 0, endRow: 0, startColumn: 0, endColumn: 1);

                Assert.Equal(A.TextAlignmentTypeValues.Center, table.GetCell(0, 0).HorizontalAlignment);
                Assert.Equal(A.TextAlignmentTypeValues.Center, table.GetCell(0, 1).HorizontalAlignment);
                Assert.Null(table.GetCell(1, 0).HorizontalAlignment);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SetCellPaddingCm_AppliesToAllCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 1, columns: 2, left: 0, top: 0, width: 4000, height: 2000);

                table.SetCellPaddingCm(0.5, 0.5, 0.5, 0.5);

                PowerPointTableCell cell = table.GetCell(0, 0);
                Assert.Equal(0.5, cell.PaddingLeftCm!.Value, 3);
                Assert.Equal(0.5, cell.PaddingTopCm!.Value, 3);
                Assert.Equal(0.5, cell.PaddingRightCm!.Value, 3);
                Assert.Equal(0.5, cell.PaddingBottomCm!.Value, 3);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SetCellBorders_AppliesToAllCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(rows: 1, columns: 1, left: 0, top: 0, width: 4000, height: 2000);

                table.SetCellBorders(TableCellBorders.All, "FF0000", widthPoints: 1);

                Assert.Equal("FF0000", table.GetCell(0, 0).BorderColor);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SetCellBorders_WithDash_WritesPresetDash() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(rows: 1, columns: 1, left: 0, top: 0, width: 4000, height: 2000);

                    table.SetCellBorders(TableCellBorders.All, "FF0000", widthPoints: 1,
                        dash: A.PresetLineDashValues.Dash);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    GraphicFrame frame = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<GraphicFrame>().First();
                    A.Table table = frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
                    A.TableCell cell = table.Elements<A.TableRow>().First().Elements<A.TableCell>().First();
                    A.PresetDash? dash = cell.TableCellProperties?.LeftBorderLineProperties?
                        .GetFirstChild<A.PresetDash>();
                    Assert.Equal(A.PresetLineDashValues.Dash, dash?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
