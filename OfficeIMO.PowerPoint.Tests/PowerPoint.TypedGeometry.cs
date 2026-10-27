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
