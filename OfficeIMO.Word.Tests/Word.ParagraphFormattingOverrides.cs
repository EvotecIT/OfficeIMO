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
            WordStyle definition = document.Styles.AddParagraphStyle("EnabledControls");
            definition.PageBreakBefore = definition.KeepWithNext = definition.KeepLinesTogether = true;
            definition.AvoidWidowAndOrphan = definition.ContextualSpacing = definition.SuppressLineNumbers = true;
            definition.SuppressAutoHyphens = definition.MirrorIndents = true;
            definition.OutlineLevel = 2;
            foreach (OnOffType property in CreateAdditionalParagraphFlags(true)) definition.Element.StyleParagraphProperties!.AddChild(property, true);
            document.AddParagraph("Before styled override");
            WordParagraph paragraph = document.AddParagraph("Styled direct off").SetStyle(definition);
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
        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(reopened.ToPdf(
            new OfficeIMO.Word.Pdf.WordPdfExportOptions { IncludePageNumbers = false, FontFamily = "Helvetica" }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Styled direct off", pdf.GetPage(1).Text);
    }

    [Fact]
    public void ParagraphFormattingOverrides_ClearRestoresInheritedFormatting() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Reset overrides");
        SetParagraphOverrides(paragraph, false);
        paragraph.OutlineLevel = 4;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(14D);
        SetParagraphOverrides(paragraph, null);
        paragraph.OutlineLevel = null;
        AssertParagraphOverrides(paragraph, null);
        Assert.Null(paragraph.OutlineLevel);
        Assert.Equal(14, paragraph.LineSpacingAfter?.ToPoints());
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
                : document.Styles.AddParagraphStyle(styleName).Element;
            style.StyleParagraphProperties ??= new StyleParagraphProperties();
            using WordDocument template = WordDocument.Create();
            WordParagraph controls = template.AddParagraph("Off controls");
            SetParagraphOverrides(controls, false);
            foreach (OnOffType property in controls._paragraph.ParagraphProperties!.Elements<OnOffType>()) {
                style.StyleParagraphProperties.AddChild(property.CloneNode(true), true);
            }
            foreach (OnOffType property in CreateAdditionalParagraphFlags(false)) style.StyleParagraphProperties.AddChild(property, true);
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
        using (WordDocument document = WordDocument.Create(path)) {
            WordStyle definition = document.Styles.AddParagraphStyle("ControlledStyle");
            definition.PageBreakBefore = definition.KeepWithNext = definition.KeepLinesTogether = state;
            definition.AvoidWidowAndOrphan = definition.ContextualSpacing = definition.SuppressLineNumbers = state;
            definition.SuppressAutoHyphens = definition.MirrorIndents = state;
            definition.OutlineLevel = 3;
            definition.FontFamily = "Arial";
            definition.LineSpacingAfter = global::OfficeIMO.OfficeLength.Emus(200L * 635);
            Assert.Empty(new OpenXmlValidator().Validate(definition.Element));
            document.AddParagraph("Uses controlled style").SetStyle(definition);
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        WordStyle loaded = reopened.Styles["ControlledStyle"];
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
        Style cleared = loaded.Element;
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
    [InlineData("Heading1")]
    [InlineData("Heading2")]
    public void NativeDoc_BuiltInStyleClearedControlsInheritInsteadOfRestoringTemplate(string styleId) {
        string path = Path.Combine(_directoryWithFiles, "ClearedControls-" + styleId + ".doc");
        using (WordDocument document = WordDocument.Create()) {
            WordStyle definition = document.Styles[styleId];
            definition.PageBreakBefore = definition.KeepWithNext = definition.KeepLinesTogether = definition.AvoidWidowAndOrphan = null;
            definition.ContextualSpacing = definition.SuppressLineNumbers = definition.SuppressAutoHyphens = definition.MirrorIndents = null;
            document.AddParagraph("Heading with inherited controls").SetStyleId(styleId);
            document.Save(path);
        }
        using WordDocument reopened = WordDocument.Load(path);
        WordStyle controls = reopened.Styles[styleId];
        Style saved = controls.Element;
        Assert.Null(controls.KeepWithNext);
        Assert.Null(controls.KeepLinesTogether);
        Assert.Null(controls.PageBreakBefore);
        Assert.Null(controls.AvoidWidowAndOrphan);
        Assert.Null(controls.ContextualSpacing);
        Assert.Null(controls.SuppressLineNumbers);
        Assert.Null(controls.SuppressAutoHyphens);
        Assert.Null(controls.MirrorIndents);
        Assert.Empty(new OpenXmlValidator().Validate(saved.StyleParagraphProperties!));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(9)]
    public void NativeDoc_OutlineAndParagraphMarkRetainValidImportedOrder(int level) {
        string path = Path.Combine(_directoryWithFiles, "OutlineAndMark-" + level + ".doc");
        using (WordDocument document = WordDocument.Create()) {
            WordParagraph paragraph = document.AddParagraph("Outline and paragraph mark");
            paragraph.OutlineLevel = level;
            paragraph._paragraph.ParagraphProperties!.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(new Bold());
            Assert.Empty(new OpenXmlValidator().Validate(paragraph._paragraph.ParagraphProperties));
            document.Save(path);
        }
        using WordDocument reopened = WordDocument.Load(path);
        WordParagraph saved = Assert.Single(reopened.Paragraphs);
        Assert.Equal(level, saved.OutlineLevel);
        Assert.NotNull(saved._paragraph.ParagraphProperties!.ParagraphMarkRunProperties?.GetFirstChild<Bold>());
        Assert.Empty(new OpenXmlValidator().Validate(saved._paragraph.ParagraphProperties));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeDoc_DefaultSectionPointsToAnEmptySepxInsteadOfTheFileHeader(bool heading) {
        string path = Path.Combine(_directoryWithFiles, "DefaultSection-" + heading + ".doc");
        using (WordDocument document = WordDocument.Create()) {
            WordParagraph paragraph = document.AddParagraph("Default section properties");
            if (heading) paragraph.Style = WordParagraphStyles.Heading1;
            document.Save(path);
        }
        byte[] data = File.ReadAllBytes(path);
        byte[] wordStream = ReadCompoundStream(data, "WordDocument");
        byte[] table = ReadCompoundStream(data, "1Table");
        int sectionTableOffset = BitConverter.ToInt32(wordStream, 202);
        int sectionTableLength = BitConverter.ToInt32(wordStream, 206);
        Assert.Equal(20, sectionTableLength); // Two CPs and one 12-byte section descriptor.
        int sepxOffset = BitConverter.ToInt32(table, sectionTableOffset + 10);
        Assert.InRange(sepxOffset, 512, wordStream.Length - 2);
        Assert.Equal(0, BitConverter.ToUInt16(wordStream, sepxOffset));
        using WordDocument reopened = WordDocument.Load(path);
        Assert.Equal("Default section properties", Assert.Single(reopened.Paragraphs).Text);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void ParagraphFormattingOverrides_RejectInvalidOutlineLevelBeforeMutation(int level) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Bounded outline");
        paragraph.OutlineLevel = 2;
        WordStyle definition = document.Styles.AddParagraphStyle("BoundedStyle");
        definition.OutlineLevel = 2;
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
