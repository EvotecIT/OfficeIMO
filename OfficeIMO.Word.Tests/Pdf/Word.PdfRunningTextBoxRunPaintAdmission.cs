using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("shadow")]
    [InlineData("outline")]
    [InlineData("emboss")]
    [InlineData("imprint")]
    public void SaveAsPdf_RunningTextBoxLegacyRunEffectRetainsWarnedContent(string effect) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        WordParagraph paragraph = box.Paragraphs.Single();
        if (effect == "shadow") paragraph.Shadow = true;
        else if (effect == "outline") paragraph.Outline = true;
        else if (effect == "emboss") paragraph.Emboss = true;
        else {
            W.Run run = box.Content!.Elements<W.Paragraph>().Single().Elements<W.Run>().Single();
            (run.RunProperties ??= new W.RunProperties()).Append(new W.Imprint());
        }
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("character")]
    [InlineData("document")]
    [InlineData("combined")]
    public void SaveAsPdf_InheritedRunningTextBoxRunEffectRetainsWarnedContent(string source) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        ConfigureRunningTextBoxInheritedShadow(document, source);
        string before = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        // Combined levels can toggle an effect; the fixed-canvas admission is deliberately
        // conservative rather than claiming that the final inherited paint is reproduced.
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_DirectOffRunningTextBoxRunEffectKeepsPreciseAdmission() {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        ConfigureRunningTextBoxInheritedShadow(document, "character");
        W.Run run = document.Sections[1].Header!.First!.TextBoxes.Single().Content!
            .Elements<W.Paragraph>().Single().Elements<W.Run>().Single();
        // The public false setter removes direct formatting; imported explicit false
        // is the distinct contract that overrides an enabled inherited effect.
        run.RunProperties!.Shadow = new W.Shadow { Val = false };
        AssertRunningTextBoxRunPaintKeepsPreciseAdmission(document);
    }

    [Fact]
    public void SaveAsPdf_NearestInheritedOffRunningTextBoxRunEffectKeepsPreciseAdmission() {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        ConfigureRunningTextBoxInheritedShadow(document, "character");
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.BasedOn { Val = "BoxRunShadow" },
            new W.StyleRunProperties(new W.Shadow { Val = false })) {
            Type = W.StyleValues.Character, StyleId = "BoxRunShadowOff"
        });
        W.Run run = document.Sections[1].Header!.First!.TextBoxes.Single().Content!
            .Elements<W.Paragraph>().Single().Elements<W.Run>().Single();
        run.RunProperties!.RunStyle!.Val = "BoxRunShadowOff";
        AssertRunningTextBoxRunPaintKeepsPreciseAdmission(document);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_UnpaintedRunningTextBoxRunEffectKeepsPreciseAdmission(bool hidden) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        W.Paragraph inner = document.Sections[1].Header!.First!.TextBoxes.Single().Content!.Elements<W.Paragraph>().Single();
        var properties = new W.RunProperties(new W.Shadow());
        if (hidden) properties.Vanish = new W.Vanish();
        // Keep the visible BOX and its terminal break intact. The other run either
        // has no glyphs or is hidden, so its effect cannot add visible paint.
        inner.PrependChild(new W.Run(properties, new W.Text(hidden ? "HIDDENINK" : string.Empty)));
        AssertRunningTextBoxRunPaintKeepsPreciseAdmission(document, "HIDDENINK");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_RunningTextBoxBodyThreeDimensionalPaintRetainsWarnedContent(bool scene) {
        using WordDocument document = CreateCompactRunningTextBoxPaintControl();
        Wps.TextBodyProperties body = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!
            .Descendants<Wps.TextBodyProperties>().Single();
        if (scene) body.AddChild(new A.Scene3DType(
            new A.Camera { Preset = A.PresetCameraValues.PerspectiveLeft },
            new A.LightRig { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top }), true);
        else body.AddChild(new A.Shape3DType(new A.BevelTop { Width = 12700L, Height = 12700L }), true);
        AssertRunningTextBoxPaintIsExplicitlyApproximated(document);
    }

    private static void ConfigureRunningTextBoxInheritedShadow(WordDocument document, string source) {
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        W.Paragraph inner = document.Sections[1].Header!.First!.TextBoxes.Single().Content!.Elements<W.Paragraph>().Single();
        W.Run run = inner.Elements<W.Run>().Single();
        if (source is "paragraph" or "combined") {
            styles.Append(new W.Style(new W.StyleRunProperties(new W.Shadow())) {
                Type = W.StyleValues.Paragraph, StyleId = "BoxParagraphShadow"
            });
            (inner.ParagraphProperties ??= new W.ParagraphProperties()).ParagraphStyleId =
                new W.ParagraphStyleId { Val = "BoxParagraphShadow" };
        }
        if (source is "character" or "combined") {
            styles.Append(new W.Style(new W.StyleRunProperties(new W.Shadow())) {
                Type = W.StyleValues.Character, StyleId = "BoxRunShadow"
            });
            (run.RunProperties ??= new W.RunProperties()).RunStyle = new W.RunStyle { Val = "BoxRunShadow" };
        }
        if (source == "document") {
            W.DocDefaults defaults = styles.DocDefaults ??= new W.DocDefaults();
            W.RunPropertiesDefault runDefaults = defaults.RunPropertiesDefault ??= new W.RunPropertiesDefault();
            W.RunPropertiesBaseStyle properties = runDefaults.RunPropertiesBaseStyle ??= new W.RunPropertiesBaseStyle();
            properties.Append(new W.Shadow());
        }
    }

    private static void AssertRunningTextBoxRunPaintKeepsPreciseAdmission(WordDocument document, string? absentText = null) {
        string[] before = AnchoredRunningStoryXml(document);
        string stylesBefore = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        var result = document.ToPdfDocumentResult(EmptyRunningOptions());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativeMixedTextBoxLayoutApproximated");
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        if (absentText != null) Assert.DoesNotContain(absentText, pdf.GetPage(2).Text);
        Assert.Equal((612D - 468D * .585D) / 2D + 7.2D,
            AnchoredRunningLetters(pdf.GetPage(2), "BOX")[0].StartBaseLine.X, 3);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(stylesBefore, document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.OuterXml);
    }
}
