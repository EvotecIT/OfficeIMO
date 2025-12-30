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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape d = slide.AddRectangle(0, 0, 500, 500);

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(0, 0, 4000, 4000), 2, 2);

                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);
                Assert.Equal(2000, a.Width);
                Assert.Equal(2000, a.Height);

                Assert.Equal(2000, b.Left);
                Assert.Equal(0, b.Top);

                Assert.Equal(0, c.Left);
                Assert.Equal(2000, c.Top);

                Assert.Equal(2000, d.Left);
                Assert.Equal(2000, d.Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 500, 500);

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(0, 0, 4200, 4300), 2, 2,
                    gutterX: 200, gutterY: 300, flow: PowerPointShapeGridFlow.ColumnMajor);

                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);

                Assert.Equal(0, b.Left);
                Assert.Equal(2300, b.Top);

                Assert.Equal(2200, c.Left);
                Assert.Equal(0, c.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeShapesInGrid_DoesNotResizeWhenDisabled() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1500);

                slide.ArrangeShapesInGrid(slide.Shapes, new PowerPointLayoutBox(0, 0, 4000, 4000), 2, 2,
                    resizeToCell: false);

                Assert.Equal(1000, a.Width);
                Assert.Equal(1500, a.Height);
                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape d = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape e = slide.AddRectangle(0, 0, 100, 100);

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(0, 0, 6000, 4000));

                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);
                Assert.Equal(2000, a.Width);
                Assert.Equal(2000, a.Height);

                Assert.Equal(2000, b.Left);
                Assert.Equal(0, b.Top);

                Assert.Equal(4000, c.Left);
                Assert.Equal(0, c.Top);

                Assert.Equal(0, d.Left);
                Assert.Equal(2000, d.Top);

                Assert.Equal(2000, e.Left);
                Assert.Equal(2000, e.Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape d = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape e = slide.AddRectangle(0, 0, 100, 100);

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(0, 0, 6000, 4000),
                    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 2 });

                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);
                Assert.Equal(3000, a.Width);
                Assert.Equal(1333, a.Height);

                Assert.Equal(3000, b.Left);
                Assert.Equal(0, b.Top);

                Assert.Equal(0, c.Left);
                Assert.Equal(1333, c.Top);

                Assert.Equal(3000, d.Left);
                Assert.Equal(1333, d.Top);

                Assert.Equal(0, e.Left);
                Assert.Equal(2666, e.Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                for (int i = 0; i < 6; i++) {
                    slide.AddRectangle(0, 0, 100, 100);
                }

                slide.ArrangeShapesInGridAuto(slide.Shapes, new PowerPointLayoutBox(0, 0, 6000, 2000),
                    new PowerPointShapeGridOptions { TargetCellAspect = 1.0 });

                PowerPointShape first = slide.Shapes[0];
                PowerPointShape fifth = slide.Shapes[4];
                PowerPointShape sixth = slide.Shapes[5];

                Assert.Equal(1200, first.Width);
                Assert.Equal(1000, first.Height);
                Assert.Equal(4800, fifth.Left);
                Assert.Equal(0, fifth.Top);
                Assert.Equal(0, sixth.Left);
                Assert.Equal(1000, sixth.Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 500, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 500, 500);
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.ArrangeShapesInGridToSlideContent(slide.Shapes, columns: 2, rows: 1,
                    marginEmus: margin);

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
                long cellWidth = content.Width / 2;

                Assert.Equal(content.Left, a.Left);
                Assert.Equal(content.Top, a.Top);
                Assert.Equal(cellWidth, a.Width);
                Assert.Equal(content.Height, a.Height);

                Assert.Equal(content.Left + cellWidth, b.Left);
                Assert.Equal(content.Top, b.Top);
                Assert.Equal(cellWidth, b.Width);
                Assert.Equal(content.Height, b.Height);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape d = slide.AddRectangle(0, 0, 100, 100);
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.ArrangeShapesInGridAutoToSlideContent(slide.Shapes, margin,
                    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 2 });

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
                long cellWidth = content.Width / 2;
                long cellHeight = content.Height / 2;

                Assert.Equal(content.Left, a.Left);
                Assert.Equal(content.Top, a.Top);

                Assert.Equal(content.Left + cellWidth, b.Left);
                Assert.Equal(content.Top, b.Top);

                Assert.Equal(content.Left, c.Left);
                Assert.Equal(content.Top + cellHeight, c.Top);

                Assert.Equal(content.Left + cellWidth, d.Left);
                Assert.Equal(content.Top + cellHeight, d.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
