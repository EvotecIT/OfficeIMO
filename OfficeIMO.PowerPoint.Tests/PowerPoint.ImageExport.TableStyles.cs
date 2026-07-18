using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsTableStyleColumnsAndCornersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();
            const string styleId = "{6A0E6B20-52C9-4C93-9DA8-000000000101}";
            AddImageExportTableStyle(slide, styleId);

            PowerPointTable table = slide.AddTablePoints(2, 3, 20, 20, 120, 60);
            table.StyleId = styleId;
            table.SetColumnWidthsPoints(40, 40, 40);
            table.SetRowHeightsPoints(30, 30);
            table.FirstRow = true;
            table.FirstColumn = true;
            table.LastColumn = true;
            table.BandedRows = false;
            table.BandedColumns = false;
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(252, 231, 243));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(219, 234, 254));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(220, 252, 231));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(226, 232, 240));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(254, 243, 199));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(252, 231, 243), image!.GetPixel(30, 30));
            Assert.Equal(OfficeColor.FromRgb(219, 234, 254), image.GetPixel(70, 30));
            Assert.Equal(OfficeColor.FromRgb(220, 252, 231), image.GetPixel(30, 65));
            Assert.Equal(OfficeColor.FromRgb(226, 232, 240), image.GetPixel(70, 65));
            Assert.Equal(OfficeColor.FromRgb(254, 243, 199), image.GetPixel(110, 65));

            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E2E8F0", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableStyleBandedColumnsThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            PowerPointSlide slide = presentation.AddSlide();
            const string styleId = "{6A0E6B20-52C9-4C93-9DA8-000000000103}";
            AddImageExportTableStyle(slide, styleId);

            PowerPointTable table = slide.AddTablePoints(2, 4, 20, 20, 120, 48);
            table.StyleId = styleId;
            table.SetColumnWidthsPoints(30, 30, 30, 30);
            table.SetRowHeightsPoints(24, 24);
            table.FirstRow = false;
            table.FirstColumn = false;
            table.LastColumn = false;
            table.BandedRows = false;
            table.BandedColumns = true;

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(199, 210, 254));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgb(254, 215, 170));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(199, 210, 254), image!.GetPixel(32, 32));
            Assert.Equal(OfficeColor.FromRgb(254, 215, 170), image.GetPixel(62, 32));
            Assert.Equal(OfficeColor.FromRgb(199, 210, 254), image.GetPixel(92, 56));
            Assert.Equal(OfficeColor.FromRgb(254, 215, 170), image.GetPixel(122, 56));

            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#C7D2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FED7AA", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsMergedTableStyleOuterBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(140, 90);
            PowerPointSlide slide = presentation.AddSlide();
            const string styleId = "{6A0E6B20-52C9-4C93-9DA8-000000000102}";
            AddImageExportMergedBorderTableStyle(slide, styleId);

            PowerPointTable table = slide.AddTablePoints(1, 2, 20, 20, 80, 36);
            table.StyleId = styleId;
            table.SetColumnWidthsPoints(40, 40);
            table.SetRowHeightsPoints(36);
            table.GetCell(0, 0).Merge = (1, 2);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);

            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line =>
                line.Shape.StrokeColor == OfficeColor.FromRgb(22, 163, 74) &&
                Math.Abs(line.Shape.StrokeWidth - 4D) < 0.000001D);
            Assert.DoesNotContain(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(250, 204, 21));

            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("#FACC15", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(22, 163, 74), image!.GetPixel(100, 38));
        }

        private static void AddImageExportTableStyle(PowerPointSlide slide, string styleId) {
            PresentationPart? presentationPart = slide.SlidePart
                .GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault();
            Assert.NotNull(presentationPart);
            PowerPointUtils.CreateTableStylesPart(presentationPart!);
            A.TableStyleList styleList = presentationPart!.TableStylesPart!.TableStyleList!;
            styleList.RemoveAllChildren<A.TableStyleEntry>();
            styleList.Append(CreateImageExportTableStyle(styleId));
        }

        private static void AddImageExportMergedBorderTableStyle(PowerPointSlide slide, string styleId) {
            PresentationPart? presentationPart = slide.SlidePart
                .GetParentParts()
                .OfType<PresentationPart>()
                .FirstOrDefault();
            Assert.NotNull(presentationPart);
            PowerPointUtils.CreateTableStylesPart(presentationPart!);
            A.TableStyleList styleList = presentationPart!.TableStylesPart!.TableStyleList!;
            styleList.RemoveAllChildren<A.TableStyleEntry>();
            string xml =
                $@"<a:tblStyle xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" styleId=""{styleId}"" styleName=""Image Export Merged Border Style"">
  <a:wholeTbl>
    <a:tcStyle>
      <a:tcBdr>
        <a:left><a:ln w=""12700"" cmpd=""sng""><a:solidFill><a:srgbClr val=""DC2626"" /></a:solidFill></a:ln></a:left>
        <a:right><a:ln w=""50800"" cmpd=""sng""><a:solidFill><a:srgbClr val=""16A34A"" /></a:solidFill></a:ln></a:right>
        <a:top><a:ln w=""12700"" cmpd=""sng""><a:solidFill><a:srgbClr val=""64748B"" /></a:solidFill></a:ln></a:top>
        <a:bottom><a:ln w=""12700"" cmpd=""sng""><a:solidFill><a:srgbClr val=""64748B"" /></a:solidFill></a:ln></a:bottom>
        <a:insideV><a:ln w=""12700"" cmpd=""sng""><a:solidFill><a:srgbClr val=""FACC15"" /></a:solidFill></a:ln></a:insideV>
      </a:tcBdr>
      <a:fill><a:solidFill><a:srgbClr val=""F8FAFC"" /></a:solidFill></a:fill>
    </a:tcStyle>
  </a:wholeTbl>
</a:tblStyle>";
            styleList.Append(new A.TableStyleEntry(xml));
        }

        private static A.TableStyleEntry CreateImageExportTableStyle(string styleId) {
            var style = new A.TableStyleEntry {
                StyleId = styleId,
                StyleName = "Image Export Table Style"
            };

            style.Append(
                new A.WholeTable(CreateTableCellStyle("E2E8F0")),
                new A.LastColumn(CreateTableCellStyle("FEF3C7")),
                new A.FirstColumn(CreateTableCellStyle("DCFCE7")),
                new A.FirstRow(CreateTableCellStyle("DBEAFE")),
                new A.NorthwestCell(CreateTableCellStyle("FCE7F3")),
                new A.Band1Vertical(CreateTableCellStyle("C7D2FE")),
                new A.Band2Vertical(CreateTableCellStyle("FED7AA")));
            return style;
        }

        private static A.TableCellStyle CreateTableCellStyle(string fillColor) =>
            new A.TableCellStyle(new A.FillProperties(new A.SolidFill(new A.RgbColorModelHex { Val = fillColor })));
    }
}
