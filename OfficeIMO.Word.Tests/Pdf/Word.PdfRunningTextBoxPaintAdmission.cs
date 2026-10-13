using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("shadow")]
    [InlineData("glow")]
    [InlineData("softEdge")]
    [InlineData("effectDag")]
    [InlineData("bevel3d")]
    [InlineData("themeEffect")]
    [InlineData("themeFill")]
    [InlineData("themeLine")]
    [InlineData("gradientOutline")]
    [InlineData("patternOutline")]
    [InlineData("customDash")]
    [InlineData("presetDash")]
    [InlineData("compoundOutline")]
    [InlineData("insetOutline")]
    [InlineData("roundCap")]
    [InlineData("roundJoin")]
    [InlineData("miterLimit")]
    [InlineData("arrow")]
    public void SaveAsPdf_UnsupportedRunningTextBoxPaintRetainsWarnedContent(string paint) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        ConfigureRunningTextBoxPaint(document, paint);
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
    }

    [Theory]
    [InlineData("border", false)]
    [InlineData("border", true)]
    [InlineData("shading", false)]
    [InlineData("shading", true)]
    [InlineData("pattern", true)]
    public void SaveAsPdf_DecoratedRunningTextBoxParagraphRetainsWarnedContent(string decoration, bool inherited) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        W.Paragraph inner = document.Sections[1].Header!.First!.TextBoxes.Single().Content!.Elements<W.Paragraph>().Single();
        OpenXmlElement authored = decoration == "border"
            ? new W.ParagraphBorders(new W.BottomBorder { Val = W.BorderValues.Single, Size = 16U, Color = "FF0000" })
            : decoration == "shading" ? new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FFFF00" }
            : new W.Shading { Val = W.ShadingPatternValues.Percent50, Color = "FF0000", Fill = "auto" };
        inner.ParagraphProperties ??= new W.ParagraphProperties();
        if (inherited) {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
                new W.StyleParagraphProperties(authored)) { Type = W.StyleValues.Paragraph, StyleId = "DecoratedBox" });
            inner.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = "DecoratedBox" };
        } else inner.ParagraphProperties.AddChild(authored, true);
        string stylesBefore = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
        Assert.Equal(stylesBefore, document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_DocumentDefaultBoxParagraphPaintRetainsWarnedContent(bool border) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults ??= new W.DocDefaults();
        W.ParagraphPropertiesDefault defaults = styles.DocDefaults.GetFirstChild<W.ParagraphPropertiesDefault>() ??
            styles.DocDefaults.AppendChild(new W.ParagraphPropertiesDefault());
        W.ParagraphPropertiesBaseStyle properties = defaults.GetFirstChild<W.ParagraphPropertiesBaseStyle>() ??
            defaults.AppendChild(new W.ParagraphPropertiesBaseStyle());
        OpenXmlElement paint = border
            ? new W.ParagraphBorders(new W.BottomBorder { Val = W.BorderValues.Single, Size = 16U, Color = "FF0000" })
            : new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FFFF00" };
        properties.AddChild(paint, true);
        string before = styles.OuterXml;
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
        Assert.Equal(before, styles.OuterXml);
    }

    [Theory]
    [InlineData("columns")]
    [InlineData("warp")]
    public void SaveAsPdf_UnsupportedRunningTextBoxBodyLayoutRetainsWarnedContent(string layout) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        Wps.TextBodyProperties body = box.Anchor!.Descendants<Wps.TextBodyProperties>().Single();
        // This run fits the full frame but exceeds one of two equal-width columns.
        // The fixed canvas would otherwise retain a single full-width line.
        WordParagraph paragraph = box.Paragraphs.Single();
        paragraph.Text = "BOX ONE TWO THREE FOUR FIVE SIX";
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 11;
        if (layout == "columns") body.ColumnCount = 2;
        else body.GetFirstChild<A.PresetTextWarp>()!.Preset = A.TextShapeValues.TextArchUp;
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document,
            new[] { "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX" });
    }

    [Theory]
    [InlineData("emptyEffects")]
    [InlineData("plainOutline")]
    [InlineData("noFillOutline")]
    public void SaveAsPdf_PlainRunningTextBoxPaintKeepsPreciseAdmission(string paint) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        ConfigureRunningTextBoxPaint(document, paint);
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal((612D - 468D * .585D) / 2D + 7.2D,
            AnchoredRunningLetters(pdf.GetPage(2), "BOX")[0].StartBaseLine.X, 3);
        Assert.Equal(paint == "plainOutline", pdf.GetPage(2).Paths.Any(path => path.IsStroked));
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    private static WordDocument CreateCompactRunningTextBoxPaintControl() {
        WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        // Keep the 51.84pt frame and its ordinary 2pt outline above the body.
        // Unsupported paint, rather than the body-overlap guard, must decide admission.
        var anchor = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!;
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
        Wps.TextBodyProperties body = anchor.Descendants<Wps.TextBodyProperties>().Single();
        body.ColumnCount = 1;
        body.AddChild(new A.PresetTextWarp(new A.AdjustValueList()) { Preset = A.TextShapeValues.TextNoShape }, true);
        return document;
    }

    private static void AssertRunningTextBoxPaintIsExplicitlyApproximated(WordDocument document, string[]? additionalTokens = null) {
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        foreach (string token in additionalTokens ?? Array.Empty<string>()) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    private static void ConfigureRunningTextBoxPaint(WordDocument document, string paint) {
        Wps.ShapeProperties shape = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!
            .Descendants<Wps.ShapeProperties>().Single();
        Wps.ShapeStyle? style = shape.Parent!.GetFirstChild<Wps.ShapeStyle>();
        if (paint.StartsWith("theme", StringComparison.Ordinal) && style == null) {
            style = new Wps.ShapeStyle(
                new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Dark1 }) { Index = A.FontCollectionIndexValues.Minor });
            ((Wps.WordprocessingShape)shape.Parent).AddChild(style, true);
        }
        switch (paint) {
            case "shadow":
                shape.AddChild(new A.EffectList(new A.OuterShadow(new A.RgbColorModelHex { Val = "FF0000" }) {
                    BlurRadius = 12700L, Distance = 25400L, Direction = 0 }), true); return;
            case "glow":
                shape.AddChild(new A.EffectList(new A.Glow(new A.RgbColorModelHex { Val = "FF0000" }) { Radius = 25400L }), true); return;
            case "softEdge":
                shape.AddChild(new A.EffectList(new A.SoftEdge { Radius = 25400L }), true); return;
            case "effectDag":
                shape.AddChild(new A.EffectDag(new A.Glow(new A.RgbColorModelHex { Val = "FF0000" }) { Radius = 25400L }), true); return;
            case "bevel3d":
                shape.AddChild(new A.Shape3DType(new A.BevelTop { Width = 12700L, Height = 12700L }), true); return;
            // The created theme's first two effect entries are empty; the third has a shadow.
            case "themeEffect": style!.GetFirstChild<A.EffectReference>()!.Index = 3U; return;
            case "themeFill": style!.GetFirstChild<A.FillReference>()!.Index = 1U; return;
            case "themeLine": style!.GetFirstChild<A.LineReference>()!.Index = 1U; return;
            case "emptyEffects": shape.AddChild(new A.EffectList(), true); return;
        }
        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) { Width = 25400 };
        switch (paint) {
            case "gradientOutline":
                outline.RemoveAllChildren();
                outline.AddChild(new A.GradientFill(new A.GradientStopList(
                    new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) { Position = 0 },
                    new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) { Position = 100000 })), true); break;
            case "patternOutline":
                outline.RemoveAllChildren();
                outline.AddChild(new A.PatternFill("<a:pattFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" prst=\"pct50\"><a:fgClr><a:srgbClr val=\"FF0000\"/></a:fgClr><a:bgClr><a:srgbClr val=\"0000FF\"/></a:bgClr></a:pattFill>"), true); break;
            case "customDash": outline.AddChild(new A.CustomDash(new A.DashStop { DashLength = 300000, SpaceLength = 100000 }), true); break;
            case "presetDash": outline.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.LargeDashDotDot }, true); break;
            case "compoundOutline": outline.CompoundLineType = A.CompoundLineValues.Double; break;
            case "insetOutline": outline.Alignment = A.PenAlignmentValues.Insert; break;
            case "roundCap": outline.CapType = A.LineCapValues.Round; break;
            case "roundJoin": outline.AddChild(new A.Round(), true); break;
            case "miterLimit": outline.AddChild(new A.Miter { Limit = 100000 }, true); break;
            case "arrow": outline.AddChild(new A.HeadEnd { Type = A.LineEndValues.Triangle }, true); break;
            case "plainOutline":
                outline.CompoundLineType = A.CompoundLineValues.Single;
                outline.Alignment = A.PenAlignmentValues.Center;
                outline.CapType = A.LineCapValues.Flat;
                outline.AddChild(new A.PresetDash { Val = A.PresetLineDashValues.Solid }, true);
                outline.AddChild(new A.HeadEnd { Type = A.LineEndValues.None }, true); break;
            case "noFillOutline":
                outline.RemoveAllChildren();
                outline.AddChild(new A.NoFill(), true);
                outline.CompoundLineType = A.CompoundLineValues.Double;
                outline.AddChild(new A.CustomDash(new A.DashStop { DashLength = 300000, SpaceLength = 100000 }), true); break;
            default: throw new ArgumentOutOfRangeException(nameof(paint), paint, null);
        }
        shape.AddChild(outline, true);
    }
}
