using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTypedGeometry {
        [Fact]
        public void SlideLengthsDriveNativeSizeAndContentRegions() {
            using var document = PowerPointDocument.Create();
            document.SlideSize.SetSize(OfficeLength.Centimeters(25.4), OfficeLength.Points(500));
            var columns = document.SlideSize.GetColumns(3, OfficeLength.Millimeters(5), OfficeLength.Points(6));
            Assert.Equal(OfficeLength.Emus(180000), columns[0].Left);
            Assert.Equal(OfficeLength.Emus(2877200), columns[0].Width);
            Assert.Throws<ArgumentOutOfRangeException>(() => document.SlideSize.GetContentBox(OfficeLength.Emus(long.MaxValue)));
            document.AddSlide();
            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var native = PresentationDocument.Open(output, false);
            Assert.Equal(9144000, native.PresentationPart!.Presentation.SlideSize!.Cx!.Value);
            Assert.Equal(6350000, native.PresentationPart.Presentation.SlideSize.Cy!.Value);
        }

        [Fact]
        public void ReturnedTableRowsAndColumnsKeepTheirNativeIdentityAfterInsertions() {
            using var document = PowerPointDocument.Create();
            var table = document.AddSlide().AddTable(2, 2);
            var row = table.AddRow(0);
            row.GetCell(0).Text = "Added row";
            var column = table.AddColumn(1);
            column.GetCell(1).Text = "Added column";
            table.AddRow(0);
            table.AddColumn(0);
            Assert.Equal(1, row.Index);
            Assert.Equal(2, column.Index);
            Assert.Equal("Added row", row.GetCell(1).Text);
            Assert.Equal("Added column", column.GetCell(2).Text);
            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var loaded = PowerPointDocument.Load(output);
            var saved = loaded.Slides[0].Tables.Single();
            Assert.Equal("Added row", saved.GetCell(1, 1).Text);
            Assert.Equal("Added column", saved.GetCell(2, 2).Text);
        }

        [Fact]
        public void MixedUnitBoundsWriteExactNativeCoordinates() {
            using var document = PowerPointDocument.Create();
            var shape = document.AddSlide().AddTextBox("Positioned");
            shape.SetPosition(OfficeLength.Inches(1), OfficeLength.Points(12));
            shape.SetSize(OfficeLength.Centimeters(2.5), OfficeLength.Pixels(96));
            Assert.Equal(OfficeLength.Emus(1814400), shape.Right);
            Assert.Equal(OfficeLength.Emus(1066800), shape.Bottom);
            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using var native = PresentationDocument.Open(output, false);
            var transform = native.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Transform2D>().Single();
            Assert.Equal(914400L, transform.Offset!.X!.Value);
            Assert.Equal(152400L, transform.Offset!.Y!.Value);
            Assert.Equal(900000L, transform.Extents!.Cx!.Value);
            Assert.Equal(914400L, transform.Extents!.Cy!.Value);
        }

        [Fact]
        public void LayoutKeepsSingleEmuInsetsAndGridRemainders() {
            var bounds = new PowerPointLayoutBox(OfficeLength.Emus(11), OfficeLength.Emus(23), OfficeLength.Emus(101), OfficeLength.Emus(61));
            var inner = bounds.Inset(OfficeLength.Emus(1));
            Assert.Equal(OfficeLength.Emus(12), inner.Left);
            Assert.Equal(OfficeLength.Emus(99), inner.Width);
            var columns = inner.SplitColumns(3, OfficeLength.Emus(1));
            Assert.Equal(OfficeLength.Emus(32), columns[0].Width);
            Assert.Equal(OfficeLength.Emus(78), columns[2].Left);
            Assert.Equal(OfficeLength.Emus(82), bounds.TakeBottom(OfficeLength.Emus(2)).Top);
            Assert.Throws<ArgumentOutOfRangeException>(() => bounds.SplitRows(2, OfficeLength.Emus(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => bounds.Inset(OfficeLength.Emus(51)));
        }
    }
}
