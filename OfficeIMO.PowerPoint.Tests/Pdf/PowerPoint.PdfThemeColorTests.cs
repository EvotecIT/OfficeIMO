using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PowerPointPdfThemeColorTests {
    [Fact]
    public void PdfThemeColors_ResolveTransformedShapeFrameAndLineColors() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        presentation.SetThemeColor(PowerPointThemeColor.Accent1, "5078A0");
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "14283C");
        presentation.SetThemeColor(PowerPointThemeColor.Accent3, "285078");
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        PowerPointAutoShape shape = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        Shape nativeShape = Assert.IsType<Shape>(shape.Element);
        SetFill(nativeShape, new A.SolidFill(new A.SchemeColor(
            new A.Shade { Val = 50000 }, new A.Alpha { Val = 50000 }) {
            Val = A.SchemeColorValues.Accent1
        }));
        PowerPointTextBox frame = slide.AddTextBoxPoints(string.Empty, 80, 10, 60, 40);
        Shape nativeFrame = Assert.IsType<Shape>(frame.Element);
        SetFill(nativeFrame, new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.Accent2
        }));
        PowerPointAutoShape line = slide.AddLinePoints(10, 90, 140, 90);
        Shape nativeLine = Assert.IsType<Shape>(line.Element);
        nativeLine.ShapeProperties!.RemoveAllChildren<A.Outline>();
        nativeLine.ShapeProperties.Append(new A.Outline(new A.SolidFill(
            new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.Accent3 })) {
            Width = 25400
        });

        Assert.Equal(OfficeColor.FromRgba(56, 86, 116, 128), shape.ResolveFillColorForExport());
        Assert.Equal(OfficeColor.FromRgb(11, 26, 41), frame.ResolveFillColorForExport());
        Assert.Equal(OfficeColor.FromRgb(26, 56, 86), line.ResolveOutlineColorForExport());
        nativeLine.ShapeProperties.GetFirstChild<A.Outline>()!.Append(new A.HeadEnd { Type = A.LineEndValues.Triangle });
        OfficeShape lineAppearance = OfficeShape.Line(0, 0, 130, 0);
        var lineDiagnostics = new List<OfficeImageExportDiagnostic>();
        PowerPointSlideImageRenderer.ApplyShapeAppearanceForExport(lineAppearance, line, lineDiagnostics);
        Assert.Empty(lineDiagnostics);
        Assert.NotNull(lineAppearance.StrokeStartMarker);
        Assert.Equal(OfficeLineMarkerKind.Triangle, lineAppearance.StrokeStartMarker!.Kind);
        // The literal editing getters remain literal; exporters use the resolved model values.
        Assert.Null(shape.FillColor);
        string pdf = PdfOperatorSearchText.From(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("0.22 0.337 0.455 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("0.043 0.102 0.161 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("0.102 0.22 0.337 RG", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.502", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfThemeColors_ApplyStyleMatrixAndReferenceTransformsInOrder() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        presentation.SetThemeColor(PowerPointThemeColor.Accent1, "5078A0");
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        A.FormatScheme format = slide.ResolvedFormatScheme!;
        A.FillStyleList fills = format.GetFirstChild<A.FillStyleList>()!;
        fills.ReplaceChild(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.PhColor
        }), fills.ChildElements[0]);
        fills.ReplaceChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }), fills.ChildElements[1]);
        A.LineStyleList lines = format.GetFirstChild<A.LineStyleList>()!;
        lines.ReplaceChild(new A.Outline(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.PhColor
        })) { Width = 12700 }, lines.ChildElements[0]);
        PowerPointAutoShape shape = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        Shape native = Assert.IsType<Shape>(shape.Element);
        SetFill(native, null);
        native.ShapeProperties!.RemoveAllChildren<A.Outline>();
        native.ShapeProperties.Append(new A.Outline { Width = 25400 });
        native.ShapeStyle = Style(1U, new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.Accent1
        }, 1U, new A.RgbColorModelHex { Val = "285078" });
        PowerPointTextBox frame = slide.AddTextBoxPoints(string.Empty, 90, 10, 60, 40);
        Shape nativeFrame = Assert.IsType<Shape>(frame.Element);
        SetFill(nativeFrame, null);
        nativeFrame.ShapeStyle = Style(2U, new A.RgbColorModelHex { Val = "0000FF" });

        Assert.Equal(OfficeColor.FromRgb(38, 61, 83), shape.ResolveFillColorForExport());
        Assert.Equal(OfficeColor.FromRgb(26, 56, 86), shape.ResolveOutlineColorForExport());
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), frame.ResolveFillColorForExport());
        string pdf = PdfOperatorSearchText.From(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("0.149 0.239 0.325 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("0.102 0.22 0.337 RG", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 0 rg", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfThemeColors_ExplicitFillAndNoFillOverrideStyles() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        PowerPointAutoShape source = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        Shape native = Assert.IsType<Shape>(source.Element);
        native.ShapeStyle = Style(1U, new A.RgbColorModelHex { Val = "FF0000" },
            1U, new A.RgbColorModelHex { Val = "00FF00" });
        SetFill(native, new A.SolidFill(new A.RgbColorModelHex(new A.Shade { Val = 50000 }) { Val = "5078A0" }));
        native.ShapeProperties!.RemoveAllChildren<A.Outline>();
        native.ShapeProperties.Append(new A.Outline(new A.NoFill()));
        Assert.Equal(OfficeColor.FromRgb(56, 86, 116), source.ResolveFillColorForExport());
        Assert.Null(source.ResolveOutlineColorForExport());
        SetFill(native, new A.NoFill());
        Assert.Null(source.ResolveFillColorForExport());
    }

    [Theory]
    [InlineData(0U)]
    [InlineData(1000U)]
    [InlineData(uint.MaxValue)]
    public void PdfThemeColors_InvalidMatrixIndicesDoNotPaintReferenceColors(uint index) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        PowerPointAutoShape source = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        Shape native = Assert.IsType<Shape>(source.Element);
        SetFill(native, null);
        native.ShapeProperties!.RemoveAllChildren<A.Outline>();
        native.ShapeStyle = Style(index, new A.RgbColorModelHex { Val = "FF0000" },
            index, new A.RgbColorModelHex { Val = "00FF00" });
        Assert.Null(source.ResolveFillColorForExport());
        Assert.Null(source.ResolveOutlineColorForExport());
    }

    [Theory]
    [InlineData(0, "102030", "0.063 0.125 0.188 rg")]
    [InlineData(1, "304050", "0.188 0.251 0.314 rg")]
    [InlineData(2, "283C50", "0.157 0.235 0.314 rg")]
    public void PdfThemeColors_UseSlideColorMapForInheritedBackgroundAndShape(int overrideKind,
        string expected, string fillOperator) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SetThemeColor(PowerPointThemeColor.Light1, "283C50");
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "102030");
        presentation.SetThemeColor(PowerPointThemeColor.Accent3, "304050");
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        SlideLayoutPart layout = slide.SlidePart.SlideLayoutPart!;
        var layoutMap = new A.OverrideColorMapping();
        layoutMap.SetAttributes(layout.SlideMasterPart!.SlideMaster.ColorMap!.GetAttributes());
        layoutMap.SetAttribute(new OpenXmlAttribute("bg1", string.Empty, "accent2"));
        layout.SlideLayout.ColorMapOverride = new ColorMapOverride(layoutMap);
        // AddSlide explicitly selects the master map. Omit that override to exercise layout inheritance.
        slide.SlidePart.Slide.ColorMapOverride = null;
        if (overrideKind == 1) {
            var slideMap = (A.OverrideColorMapping)layoutMap.CloneNode(true);
            slideMap.SetAttribute(new OpenXmlAttribute("bg1", string.Empty, "accent3"));
            slide.SlidePart.Slide.ColorMapOverride = new ColorMapOverride(slideMap);
        } else if (overrideKind == 2) {
            slide.SlidePart.Slide.ColorMapOverride = new ColorMapOverride(new A.MasterColorMapping());
        }
        layout.SlideLayout.CommonSlideData!.Background = new Background(new BackgroundProperties(
            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Background1 })));
        PowerPointAutoShape shape = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        SetFill(Assert.IsType<Shape>(shape.Element), new A.SolidFill(
            new A.SchemeColor { Val = A.SchemeColorValues.Background1 }));

        Assert.Equal(expected, shape.ResolveFillColorForExport()!.Value.ToRgbHex());
        Assert.Equal(expected, slide.GetBackground().Color);
        Assert.Contains(fillOperator, PdfOperatorSearchText.From(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })), StringComparison.Ordinal);

        layout.SlideLayout.CommonSlideData.Background = new Background(new BackgroundProperties(
            new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.Background1 }) { Position = 0 },
                new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 50000 }) {
                    Val = A.SchemeColorValues.Accent2
                }) { Position = 100000 }), new A.LinearGradientFill { Angle = 0 })));
        PowerPointSlideBackground gradient = slide.GetBackground();
        Assert.Equal(expected, gradient.GradientStartColor);
        Assert.Equal("091420", gradient.GradientEndColor);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfThemeColors_ResolveBackgroundStyleThroughThemeOverride(bool useSlideOverride) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        SlideLayoutPart layout = slide.SlidePart.SlideLayoutPart!;
        A.ThemeElements master = layout.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!;
        A.ColorScheme colors = (A.ColorScheme)master.ColorScheme!.CloneNode(true);
        colors.GetFirstChild<A.Accent2Color>()!.RemoveAllChildren();
        colors.GetFirstChild<A.Accent2Color>()!.Append(new A.RgbColorModelHex { Val = "5078A0" });
        A.FormatScheme format = (A.FormatScheme)master.FormatScheme!.CloneNode(true);
        A.BackgroundFillStyleList fills = format.GetFirstChild<A.BackgroundFillStyleList>()!;
        fills.ReplaceChild(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.PhColor
        }), fills.ChildElements[0]);
        ThemeOverridePart part = useSlideOverride ? slide.SlidePart.AddNewPart<ThemeOverridePart>()
            : layout.AddNewPart<ThemeOverridePart>();
        part.ThemeOverride = new A.ThemeOverride(colors, format);
        // A slide style reference takes precedence over a layout's explicit fill.
        layout.SlideLayout.CommonSlideData!.Background = new Background(new BackgroundProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })));
        slide.SlidePart.Slide.CommonSlideData!.Background = new Background(new BackgroundStyleReference(
            new A.SchemeColor(new A.Alpha { Val = 50000 }) { Val = A.SchemeColorValues.Accent2 }) { Index = 1001U });

        PowerPointAutoShape shape = slide.AddShapePoints(OfficePresetShapeType.Rectangle, 10, 10, 60, 40);
        Shape native = Assert.IsType<Shape>(shape.Element);
        SetFill(native, null);
        native.ShapeStyle = Style(1001U, new A.SchemeColor(new A.Alpha { Val = 50000 }) {
            Val = A.SchemeColorValues.Accent2
        });
        Assert.Equal(OfficeColor.FromRgba(56, 86, 116, 128), shape.ResolveFillColorForExport());
        Assert.Equal(OfficeColor.FromRgba(56, 86, 116, 128).ToHex(), slide.GetBackground().Color);
        string pdf = PdfOperatorSearchText.From(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("0.22 0.337 0.455 rg", pdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.502", pdf, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("run")]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("master")]
    [InlineData("fontRef")]
    [InlineData("phClr")]
    [InlineData("layout")]
    [InlineData("layoutPhClr")]
    public void PdfThemeColors_ResolveOrdinaryTextWithConsumingSlideMapAndFontReference(string source) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "112233");
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        SlideLayoutPart layout = slide.SlidePart.SlideLayoutPart!;
        A.ColorScheme colors = (A.ColorScheme)slide.ResolvedColorScheme!.CloneNode(true);
        colors.GetFirstChild<A.Accent2Color>()!.RemoveAllChildren();
        colors.GetFirstChild<A.Accent2Color>()!.Append(new A.RgbColorModelHex { Val = "5078A0" });
        slide.SlidePart.AddNewPart<ThemeOverridePart>().ThemeOverride = new A.ThemeOverride(colors);
        var map = new A.OverrideColorMapping();
        map.SetAttributes(layout.SlideMasterPart!.SlideMaster.ColorMap!.GetAttributes());
        map.SetAttribute(new OpenXmlAttribute("tx1", string.Empty, "accent2"));
        slide.SlidePart.Slide.ColorMapOverride = new ColorMapOverride(map);
        PowerPointTextBox box = slide.AddTextBoxPoints("Mapped theme text", 10, 10, 260, 70);
        box.FontName = "Helvetica";
        box.FontSize = 12;
        Shape native = Assert.IsType<Shape>(box.Element);
        PowerPointParagraph paragraph = box.Paragraphs[0];
        PowerPointTextRun run = paragraph.Runs[0];
        var mappedFill = new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
            Val = A.SchemeColorValues.Text1
        });
        if (source == "run") run.RunProperties!.Append(mappedFill);
        if (source == "paragraph") paragraph.Paragraph.ParagraphProperties = new A.ParagraphProperties(new A.DefaultRunProperties(mappedFill));
        if (source == "list") native.TextBody!.ListStyle = new A.ListStyle(new A.Level1ParagraphProperties(new A.DefaultRunProperties(mappedFill)));
        if (source == "master") {
            layout.SlideMasterPart.SlideMaster.TextStyles ??= new TextStyles();
            layout.SlideMasterPart.SlideMaster.TextStyles.OtherStyle = new OtherStyle(
                new A.Level1ParagraphProperties(new A.DefaultRunProperties(mappedFill)));
        }
        if (source is "fontRef" or "phClr") {
            native.ShapeStyle = Style(0U, new A.RgbColorModelHex { Val = "000000" });
            native.ShapeStyle.FontReference = new A.FontReference(mappedFill.ChildElements[0].CloneNode(true)) {
                Index = A.FontCollectionIndexValues.Minor
            };
        }
        if (source == "phClr") run.RunProperties!.Append(new A.SolidFill(
            new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.PhColor }));
        if (source is "layout" or "layoutPhClr") {
            box.PlaceholderType = PowerPointPlaceholderType.Body;
            box.PlaceholderIndex = 71U;
            Shape inherited = (Shape)native.CloneNode(true);
            inherited.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 71U;
            A.SolidFill inheritedFill = mappedFill;
            if (source == "layoutPhClr") {
                inheritedFill = new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) {
                    Val = A.SchemeColorValues.PhColor
                });
                inherited.ShapeStyle = Style(0U, new A.RgbColorModelHex { Val = "000000" });
                inherited.ShapeStyle.FontReference = new A.FontReference(mappedFill.ChildElements[0].CloneNode(true)) {
                    Index = A.FontCollectionIndexValues.Minor
                };
            }
            inherited.TextBody!.Elements<A.Paragraph>().First().ParagraphProperties = new A.ParagraphProperties(
                new A.DefaultRunProperties(inheritedFill));
            layout.SlideLayout.CommonSlideData!.ShapeTree!.Append(inherited);
        }
        OfficeColor expected = source is "phClr" or "layoutPhClr"
            ? OfficeColor.FromRgb(38, 61, 83) : OfficeColor.FromRgb(56, 86, 116);
        Assert.Equal(expected, PowerPointEffectiveRunStyleResolver.ResolveTextColorForExport(run, paragraph, box));
        byte[] bytes = presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        string expectedOperator = source is "phClr" or "layoutPhClr" ? "0.149 0.239 0.325 rg" : "0.22 0.337 0.455 rg";
        Assert.Contains(expectedOperator, PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
        Assert.Contains("Mapped theme text", OfficeIMO.Pdf.PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfThemeColors_UseSharedGradientStrokeAndEffectAppearance(bool textFrame) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "5078A0");
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        var map = new A.OverrideColorMapping();
        map.SetAttributes(slide.SlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.ColorMap!.GetAttributes());
        map.SetAttribute(new OpenXmlAttribute("tx1", string.Empty, "accent2"));
        slide.SlidePart.Slide.ColorMapOverride = new ColorMapOverride(map);
        PowerPointShape source = textFrame ? slide.AddTextBoxPoints(string.Empty, 20, 20, 80, 60)
            : slide.AddShapePoints(OfficePresetShapeType.Triangle, 20, 20, 80, 60);
        Shape native = Assert.IsType<Shape>(source.Element);
        SetFill(native, null);
        A.FillStyleList fills = slide.ResolvedFormatScheme!.GetFirstChild<A.FillStyleList>()!;
        fills.ReplaceChild(new A.GradientFill(new A.GradientStopList(
            new A.GradientStop(new A.SchemeColor(new A.Shade { Val = 50000 }) {
                Val = A.SchemeColorValues.PhColor
            }) { Position = 0 },
            new A.GradientStop(new A.RgbColorModelHex { Val = "A05028" }) { Position = 100000 }),
            new A.LinearGradientFill { Angle = 0 }), fills.ChildElements[0]);
        A.LineStyleList lines = slide.ResolvedFormatScheme.GetFirstChild<A.LineStyleList>()!;
        lines.ReplaceChild(new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
            new A.HeadEnd { Type = A.LineEndValues.Triangle }) { Width = 25400 }, lines.ChildElements[0]);
        native.ShapeStyle = Style(1U, new A.SchemeColor(new A.Alpha { Val = 0 }) { Val = A.SchemeColorValues.Text1 },
            1U, new A.SchemeColor { Val = A.SchemeColorValues.Text1 });
        native.ShapeProperties!.Append(new A.EffectList(
            new A.Glow(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) { Radius = 25400L },
            new A.OuterShadow(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) {
                BlurRadius = 12700L, Distance = 25400L, Direction = 0
            }));
        OfficeShape appearance = OfficeShape.Rectangle(80, 60);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        PowerPointSlideImageRenderer.ApplyShapeAppearanceForExport(appearance, source, diagnostics);
        Assert.Empty(diagnostics);
        Assert.NotNull(appearance.FillGradient);
        // A transparent placeholder stop must not suppress the matrix's fixed-color stop.
        Assert.Equal(OfficeColor.FromRgba(56, 86, 116, 0), appearance.FillGradient!.Stops[0].Color);
        Assert.Equal(OfficeColor.FromRgb(160, 80, 40), appearance.FillGradient.Stops[1].Color);
        Assert.Equal(2D, appearance.StrokeWidth);
        Assert.NotNull(appearance.Glow);
        Assert.NotNull(appearance.Shadow);
        Assert.Equal(OfficeColor.FromRgb(80, 120, 160), appearance.Glow!.Color);
        Assert.Equal(OfficeColor.FromRgb(80, 120, 160), appearance.Shadow!.Color);
        var options = new PowerPointPdfExportOptions();
        var result = presentation.ConvertToPdfWithEvidence(options);
        string pdf = PdfOperatorSearchText.From(result.ToBytes());
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == PowerPointImageExportDiagnosticCodes.UnsupportedShape);
        Assert.Contains("/ShadingType 2", pdf, StringComparison.Ordinal);
        Assert.Contains(" sh", pdf, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfThemeColors_ReportUnsupportedMatrixPatternAndPathGradient(bool pattern) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
        PowerPointTextBox frame = slide.AddTextBoxPoints(string.Empty, 20, 20, 80, 60);
        Shape native = Assert.IsType<Shape>(frame.Element);
        SetFill(native, null);
        OpenXmlElement fill;
        if (pattern) {
            fill = new A.PatternFill(new A.ForegroundColor(new A.RgbColorModelHex { Val = "FF0000" }),
                new A.BackgroundColor(new A.RgbColorModelHex { Val = "FFFFFF" }));
            fill.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "pct50"));
        } else {
            fill = new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) { Position = 100000 }),
                new A.PathGradientFill { Path = A.PathShadeValues.Rectangle });
        }
        A.FillStyleList fills = slide.ResolvedFormatScheme!.GetFirstChild<A.FillStyleList>()!;
        fills.ReplaceChild(fill, fills.ChildElements[0]);
        native.ShapeStyle = Style(1U, new A.RgbColorModelHex { Val = "FF0000" });
        var options = new PowerPointPdfExportOptions();
        var result = presentation.ConvertToPdfWithEvidence(options);
        result.ToBytes();
        Assert.Contains(result.Warnings, warning => warning.Code == PowerPointImageExportDiagnosticCodes.UnsupportedShape
            && warning.LossKind == OfficeConversionLossKind.Omission);
    }

    private static ShapeStyle Style(uint fillIndex, OpenXmlElement fillColor, uint lineIndex = 0U,
        OpenXmlElement? lineColor = null) => new ShapeStyle(
        new A.LineReference(lineColor ?? new A.RgbColorModelHex { Val = "000000" }) { Index = lineIndex },
        new A.FillReference(fillColor) { Index = fillIndex },
        new A.EffectReference { Index = 0U },
        new A.FontReference { Index = A.FontCollectionIndexValues.Minor });

    private static void SetFill(Shape shape, OpenXmlElement? fill) {
        ShapeProperties properties = shape.ShapeProperties!;
        foreach (OpenXmlElement child in properties.ChildElements.Where(child => child is A.SolidFill
                     or A.NoFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill).ToArray()) {
            child.Remove();
        }
        if (fill != null) properties.AddChild(fill, true);
    }
}
