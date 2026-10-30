using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeGridTests {
        [Fact]
        public void ArrangeShapesInGrid_RowMajor_ResizesToCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape d = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(4000),OfficeIMO.OfficeLength.Emus(4000)), 2, 2);

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(2000, a.Width.ToEmus());
                Assert.Equal(2000, a.Height.ToEmus());

                Assert.Equal(2000, b.Left.ToEmus());
                Assert.Equal(0, b.Top.ToEmus());

                Assert.Equal(0, c.Left.ToEmus());
                Assert.Equal(2000, c.Top.ToEmus());

                Assert.Equal(2000, d.Left.ToEmus());
                Assert.Equal(2000, d.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGrid_ColumnMajor_UsesGutters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(4200),OfficeIMO.OfficeLength.Emus(4300)), 2, 2,
                    gutterX: OfficeIMO.OfficeLength.Emus(200), gutterY: OfficeIMO.OfficeLength.Emus(300), flow: PowerPointShapeGridFlow.ColumnMajor);

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());

                Assert.Equal(0, b.Left.ToEmus());
                Assert.Equal(2300, b.Top.ToEmus());

                Assert.Equal(2200, c.Left.ToEmus());
                Assert.Equal(0, c.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void SplitGrid_ReturnsAlignedCellsWithIndependentGutters() {
            PowerPointLayoutBox bounds = new(OfficeIMO.OfficeLength.Emus(100),OfficeIMO.OfficeLength.Emus(200),OfficeIMO.OfficeLength.Emus(6300),OfficeIMO.OfficeLength.Emus(4400));

            PowerPointLayoutBox[,] grid = bounds.SplitGrid(2,3,rowGutter: OfficeIMO.OfficeLength.Emus(400),columnGutter: OfficeIMO.OfficeLength.Emus(300));

            Assert.Equal(2, grid.GetLength(0));
            Assert.Equal(3, grid.GetLength(1));

            Assert.Equal(100, grid[0, 0].Left.ToEmus());
            Assert.Equal(200, grid[0, 0].Top.ToEmus());
            Assert.Equal(1900, grid[0, 0].Width.ToEmus());
            Assert.Equal(2000, grid[0, 0].Height.ToEmus());

            Assert.Equal(2300, grid[0, 1].Left.ToEmus());
            Assert.Equal(4500, grid[0, 2].Left.ToEmus());
            Assert.Equal(2600, grid[1, 0].Top.ToEmus());
            Assert.Equal(grid[0, 1].Width.ToEmus(), grid[1, 1].Width.ToEmus());
            Assert.Equal(grid[0, 1].Height.ToEmus(), grid[1, 2].Height.ToEmus());
            Assert.Equal(PowerPointUnits.ToCentimeters(grid[0, 0].Right.ToEmus()), grid[0, 0].Right.ToCentimeters());
            Assert.Equal(PowerPointUnits.ToCentimeters(grid[0, 0].Bottom.ToEmus()), grid[0, 0].Bottom.ToCentimeters());
        }

        [Fact]
        public void ArrangeShapesInGrid_DoesNotResizeWhenDisabled() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1500));

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(4000),OfficeIMO.OfficeLength.Emus(4000)), 2, 2,
                    resizeToCell: false);

                Assert.Equal(1000, a.Width.ToEmus());
                Assert.Equal(1500, a.Height.ToEmus());
                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGridAuto_ComputesRowsAndColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape d = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape e = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(4000)));

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(2000, a.Width.ToEmus());
                Assert.Equal(2000, a.Height.ToEmus());

                Assert.Equal(2000, b.Left.ToEmus());
                Assert.Equal(0, b.Top.ToEmus());

                Assert.Equal(4000, c.Left.ToEmus());
                Assert.Equal(0, c.Top.ToEmus());

                Assert.Equal(0, d.Left.ToEmus());
                Assert.Equal(2000, d.Top.ToEmus());

                Assert.Equal(2000, e.Left.ToEmus());
                Assert.Equal(2000, e.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGridAuto_RespectsMinMaxColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape d = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape e = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(4000)),
                    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 2 });

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(3000, a.Width.ToEmus());
                Assert.Equal(1333, a.Height.ToEmus());

                Assert.Equal(3000, b.Left.ToEmus());
                Assert.Equal(0, b.Top.ToEmus());

                Assert.Equal(0, c.Left.ToEmus());
                Assert.Equal(1333, c.Top.ToEmus());

                Assert.Equal(3000, d.Left.ToEmus());
                Assert.Equal(1333, d.Top.ToEmus());

                Assert.Equal(0, e.Left.ToEmus());
                Assert.Equal(2666, e.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGridAuto_UsesTargetCellAspect() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                for (int i = 0; i < 6; i++) {
                    slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                }

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)),
                    new PowerPointShapeGridOptions { TargetCellAspect = 1.0 });

                PowerPointShape first = slide.Shapes[0];
                PowerPointShape fifth = slide.Shapes[4];
                PowerPointShape sixth = slide.Shapes[5];

                Assert.Equal(1200, first.Width.ToEmus());
                Assert.Equal(1000, first.Height.ToEmus());
                Assert.Equal(4800, fifth.Left.ToEmus());
                Assert.Equal(0, fifth.Top.ToEmus());
                Assert.Equal(0, sixth.Left.ToEmus());
                Assert.Equal(1000, sixth.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGridToSlideContent_RespectsMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500));
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.ArrangeShapesInGridToSlideContent(slide.Shapes, columns: 2, rows: 1,
                    margin: OfficeIMO.OfficeLength.Emus(margin));

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                long cellWidth = content.Width.ToEmus() / 2;

                Assert.Equal(content.Left.ToEmus(), a.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus(), a.Top.ToEmus());
                Assert.Equal(cellWidth, a.Width.ToEmus());
                Assert.Equal(content.Height.ToEmus(), a.Height.ToEmus());

                Assert.Equal(content.Left.ToEmus() + cellWidth, b.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus(), b.Top.ToEmus());
                Assert.Equal(cellWidth, b.Width.ToEmus());
                Assert.Equal(content.Height.ToEmus(), b.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGridAutoToSlideContent_RespectsMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape d = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.ArrangeShapesInGridAutoToSlideContent(slide.Shapes, OfficeIMO.OfficeLength.Emus(margin),
                    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 2 });

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                long cellWidth = content.Width.ToEmus() / 2;
                long cellHeight = content.Height.ToEmus() / 2;

                Assert.Equal(content.Left.ToEmus(), a.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus(), a.Top.ToEmus());

                Assert.Equal(content.Left.ToEmus() + cellWidth, b.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus(), b.Top.ToEmus());

                Assert.Equal(content.Left.ToEmus(), c.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus() + cellHeight, c.Top.ToEmus());

                Assert.Equal(content.Left.ToEmus() + cellWidth, d.Left.ToEmus());
                Assert.Equal(content.Top.ToEmus() + cellHeight, d.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        [Fact]
        public void ArrangeShapesInGridAuto_OptionsUsePhysicalGutters() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            for (int i = 0; i < 4; i++) {
                slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                    OfficeIMO.OfficeLength.Points(1), OfficeIMO.OfficeLength.Points(1));
            }
            var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(100000), OfficeIMO.OfficeLength.Emus(100000));

            slide.ArrangeShapesInGridAuto(slide.Shapes, bounds, new PowerPointShapeGridOptions {
                MinColumns = 2, MaxColumns = 2,
                GutterX = OfficeIMO.OfficeLength.Points(1),
                GutterY = OfficeIMO.OfficeLength.Points(2)
            });
            Assert.Equal(43650, slide.Shapes[0].Width.ToEmus());
            Assert.Equal(37300, slide.Shapes[0].Height.ToEmus());
            Assert.Equal(56350, slide.Shapes[1].Left.ToEmus());
            Assert.Equal(62700, slide.Shapes[2].Top.ToEmus());
        }

    }
}
