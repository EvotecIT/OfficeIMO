using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Theory]
        [InlineData("Local")]
        [InlineData("Theme")]
        [InlineData("PresetOverride")]
        public void SavedCustomDashesUseStrokeWidthAndRespectLocalPresetOverrides(string source) {
            using var presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(320, 180);
            var slide = presentation.AddSlide();
            var line = slide.AddLinePoints(20, 40, 300, 40);
            line.OutlineColor = "111111";
            line.OutlineWidthPoints = 2;
            var native = Assert.IsType<Shape>(line.Element);
            var outline = native.ShapeProperties!.GetFirstChild<A.Outline>()!;
            outline.RemoveAllChildren<A.PresetDash>();
            var custom = new A.CustomDash(
                new A.DashStop { DashLength = 100000, SpaceLength = 300000 },
                new A.DashStop { DashLength = 200000, SpaceLength = 400000 });
            if (source == "Local") outline.Append(custom);
            else {
                A.Outline themeLine = slide.ResolvedFormatScheme!.GetFirstChild<A.LineStyleList>()!.Elements<A.Outline>().First();
                themeLine.RemoveAllChildren<A.PresetDash>();
                themeLine.InsertAfter(custom, themeLine.GetFirstChild<A.SolidFill>());
                native.ShapeStyle = new ShapeStyle(
                    new A.LineReference(new A.RgbColorModelHex { Val = "111111" }) { Index = 1 },
                    new A.FillReference(new A.RgbColorModelHex { Val = "111111" }) { Index = 0 },
                    new A.EffectReference(new A.RgbColorModelHex { Val = "111111" }) { Index = 0 },
                    new A.FontReference(new A.RgbColorModelHex { Val = "111111" }) { Index = A.FontCollectionIndexValues.Minor });
                if (source == "PresetOverride") outline.Append(new A.PresetDash { Val = A.PresetLineDashValues.Solid });
            }

            using var saved = PowerPointPresentation.Load(new MemoryStream(presentation.ToBytes()));
            Assert.Empty(saved.ValidateDocument());
            var savedLine = Assert.Single(saved.Slides[0].Shapes);
            var paint = OfficeShape.Line(0, 0, 280, 0);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            PowerPointSlideImageRenderer.ApplyShapeAppearanceForExport(paint, savedLine, diagnostics);
            Assert.Empty(diagnostics);
            if (source == "PresetOverride") Assert.Empty(paint.StrokeDashArray);
            else Assert.Equal(new[] { 2D, 6D, 4D, 8D }, paint.StrokeDashArray);

            string pdf = PdfOperatorSearchText.From(saved.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            if (source == "PresetOverride") Assert.DoesNotContain("[2 6 4 8] 0 d", pdf, StringComparison.Ordinal);
            else Assert.Contains("[2 6 4 8] 0 d", pdf, StringComparison.Ordinal);
        }
    }
}
