using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void ParagraphFormattingOverrides_RoundTripOnOffAndInherited(WordFileFormat format) {
        string path = Path.Combine(_directoryWithFiles, "FormattingOverrides." + (format == WordFileFormat.Doc ? "doc" : "docx"));
        using (WordDocument document = WordDocument.Create()) {
            foreach (bool? state in new bool?[] { true, false, null }) {
                WordParagraph paragraph = document.AddParagraph("Formatting " + (state?.ToString() ?? "inherited"));
                SetParagraphOverrides(paragraph, state);
                paragraph.OutlineLevel = state == null ? null : state == true ? 2 : 9;
                Assert.Empty(new OpenXmlValidator().Validate(paragraph._paragraph.ParagraphProperties!));
            }
            document.Save(path);
        }

        using WordDocument reopened = WordDocument.Load(path);
        Assert.Equal(3, reopened.Paragraphs.Count);
        AssertParagraphOverrides(reopened.Paragraphs[0], true);
        AssertParagraphOverrides(reopened.Paragraphs[1], false);
        AssertParagraphOverrides(reopened.Paragraphs[2], null);
        Assert.Equal(2, reopened.Paragraphs[0].OutlineLevel);
        Assert.Equal(9, reopened.Paragraphs[1].OutlineLevel);
    }

    [Fact]
    public void NativeDoc_ParagraphOffOverridesEnabledStyle() {
        string path = Path.Combine(_directoryWithFiles, "StyledOffOverrides.doc");
        using (WordDocument document = WordDocument.Create()) {
            var definition = new WordParagraphStyleDefinition("EnabledControls") {
                BasedOnStyleId = "Normal", PageBreakBefore = true, KeepWithNext = true,
                KeepLinesTogether = true, AvoidWidowAndOrphan = true, ContextualSpacing = true,
                SuppressLineNumbers = true, SuppressAutoHyphens = true, MirrorIndents = true, OutlineLevel = 2
            };
            Style style = definition.ToOpenXml();
            foreach (OnOffType property in CreateAdditionalParagraphFlags(true)) style.StyleParagraphProperties!.AddChild(property, true);
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            document.AddParagraph("Before styled override");
            WordParagraph paragraph = document.AddParagraph("Styled direct off").SetStyleId(definition.StyleId);
            SetParagraphOverrides(paragraph, false);
            paragraph.OutlineLevel = 9;
            foreach (OnOffType property in CreateAdditionalParagraphFlags(false)) paragraph._paragraph.ParagraphProperties!.AddChild(property, true);
            document.Save(path);
        }

        using WordDocument reopened = WordDocument.Load(path);
        WordParagraph target = reopened.Paragraphs.Single(p => p.Text == "Styled direct off");
        AssertParagraphOverrides(target, false);
        Assert.Equal(9, target.OutlineLevel);
        foreach (OnOffType property in CreateAdditionalParagraphFlags(false)) {
            OnOffType saved = Assert.Single(target._paragraph.ParagraphProperties!.Elements<OnOffType>(), p => p.GetType() == property.GetType());
            Assert.False(saved.Val!.Value);
        }
        Style enabled = reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(s => s.StyleName?.Val?.Value == "EnabledControls");
        Assert.True(enabled.StyleParagraphProperties?.PageBreakBefore?.Val?.Value ?? enabled.StyleParagraphProperties?.PageBreakBefore != null);
        Assert.Equal(2, enabled.StyleParagraphProperties?.OutlineLevel?.Val?.Value);
        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(reopened.ToPdfBytes(
            new OfficeIMO.Word.Pdf.WordToPdfOptions { IncludePageNumbers = false, FontFamily = "Helvetica" }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Styled direct off", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ParagraphFormattingOverrides_ClearRestoresInheritedFormatting() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Reset overrides");
        SetParagraphOverrides(paragraph, false);
        paragraph.OutlineLevel = 4;
        paragraph.LineSpacingAfterPoints = 14;
        SetParagraphOverrides(paragraph, null);
        paragraph.OutlineLevel = null;
        AssertParagraphOverrides(paragraph, null);
        Assert.Null(paragraph.OutlineLevel);
        Assert.Equal(14, paragraph.LineSpacingAfterPoints);
        Assert.Empty(new OpenXmlValidator().Validate(paragraph._paragraph.ParagraphProperties!));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDoc_StyleRetainsExplicitOffControls(bool builtInStyle) {
        string path = Path.Combine(_directoryWithFiles, "StyleOffControls-" + builtInStyle + ".doc");
        string styleName = builtInStyle ? "Normal" : "DisabledControls";
        using (WordDocument document = WordDocument.Create()) {
            var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style style = builtInStyle
                ? styles.Elements<Style>().Single(s => s.StyleId?.Value == "Normal")
                : new WordParagraphStyleDefinition(styleName) { BasedOnStyleId = "Normal" }.ToOpenXml();
            style.StyleParagraphProperties ??= new StyleParagraphProperties();
            using WordDocument template = WordDocument.Create();
            WordParagraph controls = template.AddParagraph("Off controls");
            SetParagraphOverrides(controls, false);
            foreach (OnOffType property in controls._paragraph.ParagraphProperties!.Elements<OnOffType>()) {
                style.StyleParagraphProperties.AddChild(property.CloneNode(true), true);
            }
            foreach (OnOffType property in CreateAdditionalParagraphFlags(false)) style.StyleParagraphProperties.AddChild(property, true);
            if (!builtInStyle) styles.Append(style);
            document.AddParagraph("Paragraph using off controls").SetStyleId(style.StyleId!.Value!);
            document.Save(path);
        }
        using WordDocument reopened = WordDocument.Load(path);
        Style saved = reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(s => s.StyleName?.Val?.Value == styleName);
        Assert.Empty(new OpenXmlValidator().Validate(saved.StyleParagraphProperties!));
        Assert.Equal(15, saved.StyleParagraphProperties!.Elements<OnOffType>().Count());
        Assert.All(saved.StyleParagraphProperties.Elements<OnOffType>(), p => Assert.False(p.Val!.Value));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [InlineData(null)]
    public void ParagraphStyleFormattingOverrides_RoundTripAndClear(bool? state) {
        string path = Path.Combine(_directoryWithFiles, "StyleFormatting.docx");
        var definition = new WordParagraphStyleDefinition("ControlledStyle") {
            BasedOnStyleId = "Normal", PageBreakBefore = state, KeepWithNext = state,
            KeepLinesTogether = state, AvoidWidowAndOrphan = state, ContextualSpacing = state,
            SuppressLineNumbers = state, SuppressAutoHyphens = state, MirrorIndents = state,
            OutlineLevel = 3, FontName = "Arial", SpacingAfterTwips = 200
        };
        using (WordDocument document = WordDocument.Create(path)) {
            Style style = definition.ToOpenXml();
            Assert.Empty(new OpenXmlValidator().Validate(style));
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
            document.AddParagraph("Uses controlled style").SetStyleId(definition.StyleId);
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        Style saved = reopened._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>().Single(s => s.StyleId == definition.StyleId);
        var loaded = new WordParagraphStyleDefinition(saved);
        Assert.Equal(state, loaded.PageBreakBefore);
        Assert.Equal(state, loaded.KeepWithNext);
        Assert.Equal(state, loaded.KeepLinesTogether);
        Assert.Equal(state, loaded.AvoidWidowAndOrphan);
        Assert.Equal(state, loaded.ContextualSpacing);
        Assert.Equal(state, loaded.SuppressLineNumbers);
        Assert.Equal(state, loaded.SuppressAutoHyphens);
        Assert.Equal(state, loaded.MirrorIndents);
        Assert.Equal(3, loaded.OutlineLevel);
        loaded.PageBreakBefore = loaded.KeepWithNext = loaded.KeepLinesTogether = loaded.AvoidWidowAndOrphan = null;
        loaded.ContextualSpacing = loaded.SuppressLineNumbers = loaded.SuppressAutoHyphens = loaded.MirrorIndents = null;
        loaded.OutlineLevel = null;
        Style cleared = loaded.ToOpenXml();
        Assert.Empty(new OpenXmlValidator().Validate(cleared));
        Assert.Null(cleared.StyleParagraphProperties?.PageBreakBefore);
        Assert.Null(cleared.StyleParagraphProperties?.KeepNext);
        Assert.Null(cleared.StyleParagraphProperties?.KeepLines);
        Assert.Null(cleared.StyleParagraphProperties?.WidowControl);
        Assert.Null(cleared.StyleParagraphProperties?.ContextualSpacing);
        Assert.Null(cleared.StyleParagraphProperties?.SuppressLineNumbers);
        Assert.Null(cleared.StyleParagraphProperties?.SuppressAutoHyphens);
        Assert.Null(cleared.StyleParagraphProperties?.MirrorIndents);
        Assert.Null(cleared.StyleParagraphProperties?.OutlineLevel);
        Assert.Equal("200", cleared.StyleParagraphProperties?.SpacingBetweenLines?.After?.Value);
        Assert.Equal("Arial", cleared.StyleRunProperties?.RunFonts?.Ascii?.Value);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void ParagraphFormattingOverrides_RejectInvalidOutlineLevelBeforeMutation(int level) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Bounded outline");
        paragraph.OutlineLevel = 2;
        var definition = new WordParagraphStyleDefinition("BoundedStyle") { OutlineLevel = 2 };
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.OutlineLevel = level);
        Assert.Throws<ArgumentOutOfRangeException>(() => definition.OutlineLevel = level);
        Assert.Equal(2, paragraph.OutlineLevel);
        Assert.Equal(2, definition.OutlineLevel);
    }

    private static void SetParagraphOverrides(WordParagraph paragraph, bool? value) {
        paragraph.PageBreakBeforeOverride = value;
        paragraph.KeepWithNextOverride = value;
        paragraph.KeepLinesTogetherOverride = value;
        paragraph.AvoidWidowAndOrphanOverride = value;
        paragraph.ContextualSpacing = value;
        paragraph.SuppressLineNumbers = value;
        paragraph.SuppressAutoHyphens = value;
        paragraph.MirrorIndents = value;
    }

    private static void AssertParagraphOverrides(WordParagraph paragraph, bool? value) {
        Assert.Equal(value, paragraph.PageBreakBeforeOverride);
        Assert.Equal(value, paragraph.KeepWithNextOverride);
        Assert.Equal(value, paragraph.KeepLinesTogetherOverride);
        Assert.Equal(value, paragraph.AvoidWidowAndOrphanOverride);
        Assert.Equal(value, paragraph.ContextualSpacing);
        Assert.Equal(value, paragraph.SuppressLineNumbers);
        Assert.Equal(value, paragraph.SuppressAutoHyphens);
        Assert.Equal(value, paragraph.MirrorIndents);
    }

    private static OnOffType[] CreateAdditionalParagraphFlags(bool value) => new OnOffType[] {
        new Kinsoku { Val = value }, new WordWrap { Val = value },
        new OverflowPunctuation { Val = value }, new TopLinePunctuation { Val = value },
        new AutoSpaceDE { Val = value }, new AutoSpaceDN { Val = value }, new BiDi { Val = value }
    };
}
