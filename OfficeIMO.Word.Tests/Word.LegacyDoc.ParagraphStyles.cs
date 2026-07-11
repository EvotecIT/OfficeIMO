using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenMcdf;
using Xunit;
using Version = OpenMcdf.Version;
using StorageModeFlags = OpenMcdf.StorageModeFlags;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyles() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithHeadingStyles();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(3, paragraphs.Length);
            Assert.Equal("Heading One", paragraphs[0].Text);
            Assert.Equal(WordParagraphStyles.Heading1, paragraphs[0].Style);
            Assert.Equal("Heading Two", paragraphs[1].Text);
            Assert.Equal(WordParagraphStyles.Heading2, paragraphs[1].Style);
            Assert.Equal("Body", paragraphs[2].Text);
            Assert.Null(paragraphs[2].Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            WordParagraph[] convertedParagraphs = converted.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(WordParagraphStyles.Heading1, convertedParagraphs[0].Style);
            Assert.Equal(WordParagraphStyles.Heading2, convertedParagraphs[1].Style);
            Assert.Null(convertedParagraphs[2].Style);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph[] paragraphs = result.Document.Paragraphs
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToArray();

            Assert.Equal(2, paragraphs.Length);
            Assert.Equal("Styled Custom", paragraphs[0].Text);
            Assert.Equal(WordParagraphStyles.Custom, paragraphs[0].Style);
            Assert.Equal("LegacyDocCustomBody", paragraphs[0].StyleId);
            Assert.Equal("Body", paragraphs[1].Text);
            Assert.Null(paragraphs[1].Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            WordParagraph convertedParagraph = converted.Paragraphs
                .First(paragraph => paragraph.Text == "Styled Custom");
            Assert.Equal(WordParagraphStyles.Custom, convertedParagraph.Style);
            Assert.Equal("LegacyDocCustomBody", convertedParagraph.StyleId);

            DocumentFormat.OpenXml.Wordprocessing.Style? customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<DocumentFormat.OpenXml.Wordprocessing.Style>()
                .FirstOrDefault(style => style.StyleId?.Value == "LegacyDocCustomBody");
            Assert.NotNull(customStyle);
            Assert.Equal("Custom Body", customStyle!.StyleName?.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleFormattingFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleFormatting();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Custom Formatting");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomFormattedBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomFormattedBody");

            Assert.Equal("Custom Formatted Body", customStyle.StyleName?.Val?.Value);
            Assert.Equal("Heading1", customStyle.BasedOn?.Val?.Value);
            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.Equal(JustificationValues.Center, paragraphProperties.GetFirstChild<Justification>()?.Val?.Value);
            Assert.Equal("240", paragraphProperties.GetFirstChild<SpacingBetweenLines>()?.After?.Value);
            Assert.Equal("720", paragraphProperties.GetFirstChild<Indentation>()?.Left?.Value);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(customStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(runProperties.GetFirstChild<Bold>());
            Assert.Equal("28", runProperties.GetFirstChild<FontSize>()?.Val?.Value);
            Assert.Equal("ff0000", runProperties.GetFirstChild<Color>()?.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_PreservesCustomParagraphStyleInheritance() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleInheritance();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Custom Child");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomChild", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Styles styles = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style baseStyle = styles
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomBase");
            Assert.Equal("Custom Base", baseStyle.StyleName?.Val?.Value);
            StyleParagraphProperties baseParagraphProperties = Assert.IsType<StyleParagraphProperties>(baseStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.Equal(JustificationValues.Center, baseParagraphProperties.GetFirstChild<Justification>()?.Val?.Value);
            StyleRunProperties baseRunProperties = Assert.IsType<StyleRunProperties>(baseStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(baseRunProperties.GetFirstChild<Bold>());
            Assert.NotNull(baseRunProperties.GetFirstChild<BoldComplexScript>());

            Style childStyle = styles
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomChild");
            Assert.Equal("Custom Child", childStyle.StyleName?.Val?.Value);
            Assert.Equal("LegacyDocCustomBase", childStyle.BasedOn?.Val?.Value);
            StyleRunProperties childRunProperties = Assert.IsType<StyleRunProperties>(childStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(childRunProperties.GetFirstChild<Italic>());
            Assert.NotNull(childRunProperties.GetFirstChild<ItalicComplexScript>());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleNextStyle() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleNextStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Next");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomNextSource", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Styles styles = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style sourceStyle = styles
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomNextSource");
            Assert.Equal("LegacyDocCustomNextTarget", sourceStyle.GetFirstChild<NextParagraphStyle>()?.Val?.Value);

            Style targetStyle = styles
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomNextTarget");
            Assert.Equal("Custom Next Target", targetStyle.StyleName?.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleNextStyle() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleNextStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Next");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Styles styles = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style headingStyle = styles
                .OfType<Style>()
                .First(style => style.StyleId?.Value == WordParagraphStyles.Heading1.ToStringStyle());
            Assert.Equal(
                WordParagraphStyles.Heading2.ToStringStyle(),
                headingStyle.GetFirstChild<NextParagraphStyle>()?.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStylePaginationFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStylePaginationFlags();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Pagination");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomPaginationBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomPaginationBody");

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.NotNull(paragraphProperties.GetFirstChild<KeepLines>());
            Assert.NotNull(paragraphProperties.GetFirstChild<KeepNext>());
            Assert.NotNull(paragraphProperties.GetFirstChild<PageBreakBefore>());
            Assert.NotNull(paragraphProperties.GetFirstChild<WidowControl>());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStylePaginationFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStylePaginationFlags();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Pagination");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.NotNull(paragraphProperties.GetFirstChild<KeepLines>());
            Assert.NotNull(paragraphProperties.GetFirstChild<KeepNext>());
            Assert.NotNull(paragraphProperties.GetFirstChild<PageBreakBefore>());
            Assert.NotNull(paragraphProperties.GetFirstChild<WidowControl>());
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleShadingFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleShading();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Shading");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomShadedBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomShadedBody");

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.GetFirstChild<StyleParagraphProperties>());
            Shading shading = Assert.IsType<Shading>(paragraphProperties.GetFirstChild<Shading>());
            Assert.Equal(ShadingPatternValues.Clear, shading.Val!.Value);
            Assert.Equal("auto", shading.Color!.Value);
            Assert.Equal("ff0000", shading.Fill!.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleShadingFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleShading();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Shading");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyle.GetFirstChild<StyleParagraphProperties>());
            Shading shading = Assert.IsType<Shading>(paragraphProperties.GetFirstChild<Shading>());
            Assert.Equal(ShadingPatternValues.Clear, shading.Val!.Value);
            Assert.Equal("auto", shading.Color!.Value);
            Assert.Equal("ff0000", shading.Fill!.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleLayoutFlagsFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleLayoutFlags();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Layout Flags");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomLayoutFlagsBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomLayoutFlagsBody");

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.NotNull(paragraphProperties.GetFirstChild<SuppressLineNumbers>());
            Assert.NotNull(paragraphProperties.GetFirstChild<SuppressAutoHyphens>());
            Assert.NotNull(paragraphProperties.GetFirstChild<ContextualSpacing>());
            Assert.NotNull(paragraphProperties.GetFirstChild<MirrorIndents>());
            Assert.NotNull(paragraphProperties.GetFirstChild<Kinsoku>());
            Assert.NotNull(paragraphProperties.GetFirstChild<WordWrap>());
            Assert.NotNull(paragraphProperties.GetFirstChild<OverflowPunctuation>());
            Assert.NotNull(paragraphProperties.GetFirstChild<TopLinePunctuation>());
            Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDE>());
            Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDN>());
            Assert.NotNull(paragraphProperties.GetFirstChild<BiDi>());
            TextAlignment textAlignment = Assert.IsType<TextAlignment>(paragraphProperties.GetFirstChild<TextAlignment>());
            Assert.Equal(VerticalTextAlignmentValues.Center, textAlignment.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleLayoutFlagsFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleLayoutFlags();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Layout Flags");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);

            StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyle.GetFirstChild<StyleParagraphProperties>());
            Assert.NotNull(paragraphProperties.GetFirstChild<SuppressLineNumbers>());
            Assert.NotNull(paragraphProperties.GetFirstChild<SuppressAutoHyphens>());
            Assert.NotNull(paragraphProperties.GetFirstChild<ContextualSpacing>());
            Assert.NotNull(paragraphProperties.GetFirstChild<MirrorIndents>());
            Assert.NotNull(paragraphProperties.GetFirstChild<Kinsoku>());
            Assert.NotNull(paragraphProperties.GetFirstChild<WordWrap>());
            Assert.NotNull(paragraphProperties.GetFirstChild<OverflowPunctuation>());
            Assert.NotNull(paragraphProperties.GetFirstChild<TopLinePunctuation>());
            Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDE>());
            Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDN>());
            Assert.NotNull(paragraphProperties.GetFirstChild<BiDi>());
            TextAlignment textAlignment = Assert.IsType<TextAlignment>(paragraphProperties.GetFirstChild<TextAlignment>());
            Assert.Equal(VerticalTextAlignmentValues.Center, textAlignment.Val?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleFontFamilyFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleFontFamily();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Font Family");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomFontBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomFontBody");
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(customStyle.GetFirstChild<StyleRunProperties>());
            RunFonts runFonts = Assert.IsType<RunFonts>(runProperties.GetFirstChild<RunFonts>());
            Assert.Equal("Courier New", runFonts.Ascii?.Value);
            Assert.Equal("Courier New", runFonts.HighAnsi?.Value);
            Assert.Equal("Courier New", runFonts.ComplexScript?.Value);
            Assert.Equal("Courier New", runFonts.EastAsia?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleFontFamilyFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleFontFamily();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Font Family");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(headingStyle.GetFirstChild<StyleRunProperties>());
            RunFonts runFonts = Assert.IsType<RunFonts>(runProperties.GetFirstChild<RunFonts>());
            Assert.Equal("Courier New", runFonts.Ascii?.Value);
            Assert.Equal("Courier New", runFonts.HighAnsi?.Value);
            Assert.Equal("Courier New", runFonts.ComplexScript?.Value);
            Assert.Equal("Courier New", runFonts.EastAsia?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleLanguageFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleLanguage();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Language");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomLanguageBody", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomLanguageBody");
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(customStyle.GetFirstChild<StyleRunProperties>());
            Languages languages = Assert.IsType<Languages>(runProperties.GetFirstChild<Languages>());
            Assert.Equal("pl-PL", languages.Val?.Value);
            Assert.Equal("ja-JP", languages.EastAsia?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleLanguageFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleLanguage();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Language");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(headingStyle.GetFirstChild<StyleRunProperties>());
            Languages languages = Assert.IsType<Languages>(runProperties.GetFirstChild<Languages>());
            Assert.Equal("pl-PL", languages.Val?.Value);
            Assert.Equal("ja-JP", languages.EastAsia?.Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsCustomParagraphStyleItalicUnderlineStrikeVerticalAndHighlightFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithCustomParagraphStyleItalicUnderlineStrikeVerticalAndHighlight();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Styled Italic Underline Strike Super Mark");
            Assert.Equal(WordParagraphStyles.Custom, paragraph.Style);
            Assert.Equal("LegacyDocCustomItalicUnderlineStrikeSuperMark", paragraph.StyleId);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style customStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == "LegacyDocCustomItalicUnderlineStrikeSuperMark");
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(customStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(runProperties.GetFirstChild<Italic>());
            Assert.Equal(UnderlineValues.Single, runProperties.GetFirstChild<Underline>()?.Val?.Value);
            Assert.NotNull(runProperties.GetFirstChild<Strike>());
            Assert.Equal(VerticalPositionValues.Superscript, runProperties.GetFirstChild<VerticalTextAlignment>()?.Val?.Value);
            DocumentFormat.OpenXml.OpenXmlElement highlight = Assert.Single(runProperties.ChildElements, element => element.LocalName == "highlight");
            Assert.Equal("yellow", highlight.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
        }

        [Fact]
        public void LegacyDoc_LoadLegacyDocWithReport_ProjectsBuiltInParagraphStyleItalicUnderlineStrikeVerticalAndHighlightFromStyleSheet() {
            byte[] docBytes = LegacyDocParagraphStyleFixture.CreateDocWithBuiltInParagraphStyleItalicUnderlineStrikeVerticalAndHighlight();
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport(new MemoryStream(docBytes));

            result.EnsureNoImportErrors();
            WordParagraph paragraph = Assert.Single(
                result.Document.Paragraphs,
                item => item.Text == "Heading Italic Underline Strike Super Mark");
            Assert.Equal(WordParagraphStyles.Heading1, paragraph.Style);

            using WordDocument converted = WordDocument.Load(new MemoryStream(result.Document.ToDocx()));
            Style headingStyle = converted._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!
                .OfType<Style>()
                .First(style => style.StyleId?.Value == headingStyleId);
            StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(headingStyle.GetFirstChild<StyleRunProperties>());
            Assert.NotNull(runProperties.GetFirstChild<Italic>());
            Assert.Equal(UnderlineValues.Single, runProperties.GetFirstChild<Underline>()?.Val?.Value);
            Assert.NotNull(runProperties.GetFirstChild<Strike>());
            Assert.Equal(VerticalPositionValues.Superscript, runProperties.GetFirstChild<VerticalTextAlignment>()?.Val?.Value);
            DocumentFormat.OpenXml.OpenXmlElement highlight = Assert.Single(runProperties.ChildElements, element => element.LocalName == "highlight");
            Assert.Equal("yellow", highlight.GetAttribute("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value);
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphStylesAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Heading One").SetStyle(WordParagraphStyles.Heading1);
                    document.AddParagraph("Heading Two").SetStyle(WordParagraphStyles.Heading2);
                    document.AddParagraph("Body");

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                WordParagraph[] paragraphs = reloaded.Paragraphs
                    .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                    .ToArray();

                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                Assert.Equal(3, paragraphs.Length);
                Assert.Equal("Heading One", paragraphs[0].Text);
                Assert.Equal(WordParagraphStyles.Heading1, paragraphs[0].Style);
                Assert.Equal("Heading Two", paragraphs[1].Text);
                Assert.Equal(WordParagraphStyles.Heading2, paragraphs[1].Style);
                Assert.Equal("Body", paragraphs[2].Text);
                Assert.Null(paragraphs[2].Style);
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocParagraphStyleNextStyleAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string sourceStyleId = "NativeDocNextSource";
            const string targetStyleId = "NativeDocNextTarget";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    var sourceStyle = new Style { Type = StyleValues.Paragraph, StyleId = sourceStyleId, CustomStyle = true };
                    sourceStyle.Append(new StyleName { Val = "Native DOC Next Source" });
                    sourceStyle.Append(new BasedOn { Val = WordParagraphStyles.Normal.ToStringStyle() });
                    sourceStyle.Append(new NextParagraphStyle { Val = targetStyleId });
                    styles.Append(sourceStyle);

                    var targetStyle = new Style { Type = StyleValues.Paragraph, StyleId = targetStyleId, CustomStyle = true };
                    targetStyle.Append(new StyleName { Val = "Native DOC Next Target" });
                    targetStyle.Append(new BasedOn { Val = WordParagraphStyles.Normal.ToStringStyle() });
                    targetStyle.Append(new StyleRunProperties(new Italic()));
                    styles.Append(targetStyle);

                    document.AddParagraph("Native next style source").SetStyleId(sourceStyleId);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Native next style source", paragraph.Text);
                Assert.Equal("LegacyDocNativeDOCNextSource", paragraph.StyleId);

                Styles reloadedStyles = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style sourceStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCNextSource");
                Assert.Equal("LegacyDocNativeDOCNextTarget", sourceStyleAfterReload.GetFirstChild<NextParagraphStyle>()?.Val?.Value);

                Style targetStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCNextTarget");
                Assert.Equal("Native DOC Next Target", targetStyleAfterReload.StyleName?.Val?.Value);
                StyleRunProperties runProperties = Assert.IsType<StyleRunProperties>(targetStyleAfterReload.StyleRunProperties);
                Assert.NotNull(runProperties.GetFirstChild<Italic>());
                Assert.NotNull(runProperties.GetFirstChild<ItalicComplexScript>());
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomParagraphStyleInheritanceAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string baseStyleId = "NativeDocBase";
            const string childStyleId = "NativeDocChild";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    var baseStyle = new Style { Type = StyleValues.Paragraph, StyleId = baseStyleId, CustomStyle = true };
                    baseStyle.Append(new StyleName { Val = "Native DOC Base" });
                    baseStyle.Append(new BasedOn { Val = WordParagraphStyles.Normal.ToStringStyle() });
                    baseStyle.Append(new StyleParagraphProperties(new Justification { Val = JustificationValues.Center }));
                    baseStyle.Append(new StyleRunProperties(new Bold(), new BoldComplexScript()));
                    styles.Append(baseStyle);

                    var childStyle = new Style { Type = StyleValues.Paragraph, StyleId = childStyleId, CustomStyle = true };
                    childStyle.Append(new StyleName { Val = "Native DOC Child" });
                    childStyle.Append(new BasedOn { Val = baseStyleId });
                    childStyle.Append(new StyleRunProperties(new Italic(), new ItalicComplexScript()));
                    styles.Append(childStyle);

                    document.AddParagraph("Native custom child").SetStyleId(childStyleId);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Native custom child", paragraph.Text);
                Assert.Equal("LegacyDocNativeDOCChild", paragraph.StyleId);

                Styles reloadedStyles = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style baseStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCBase");
                Assert.Equal("Native DOC Base", baseStyleAfterReload.StyleName?.Val?.Value);
                StyleParagraphProperties baseParagraphProperties = Assert.IsType<StyleParagraphProperties>(baseStyleAfterReload.StyleParagraphProperties);
                Assert.Equal(JustificationValues.Center, baseParagraphProperties.GetFirstChild<Justification>()?.Val?.Value);
                StyleRunProperties baseRunProperties = Assert.IsType<StyleRunProperties>(baseStyleAfterReload.StyleRunProperties);
                Assert.NotNull(baseRunProperties.GetFirstChild<Bold>());
                Assert.NotNull(baseRunProperties.GetFirstChild<BoldComplexScript>());

                Style childStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCChild");
                Assert.Equal("Native DOC Child", childStyleAfterReload.StyleName?.Val?.Value);
                Assert.Equal("LegacyDocNativeDOCBase", childStyleAfterReload.BasedOn?.Val?.Value);
                StyleRunProperties childRunProperties = Assert.IsType<StyleRunProperties>(childStyleAfterReload.StyleRunProperties);
                Assert.NotNull(childRunProperties.GetFirstChild<Italic>());
                Assert.NotNull(childRunProperties.GetFirstChild<ItalicComplexScript>());
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomParagraphStyleNumberingAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string styleId = "NativeDocNumberedStyle";
            const string projectedStyleId = "LegacyDocNativeDOCNumberedStyle";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId, CustomStyle = true };
                    style.Append(new StyleName { Val = "Native DOC Numbered Style" });
                    style.Append(new BasedOn { Val = WordParagraphStyles.Normal.ToStringStyle() });
                    style.Append(new StyleParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 2 },
                            new NumberingId { Val = 9 })));
                    document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
                    document.AddParagraph("Styled numbered paragraph").SetStyleId(styleId);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Styled numbered paragraph", paragraph.Text);
                Assert.Equal(projectedStyleId, paragraph.StyleId);

                Styles styles = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style customStyle = Assert.Single(styles.Elements<Style>(), styleDefinition => styleDefinition.StyleId == projectedStyleId);
                StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(customStyle.StyleParagraphProperties);
                NumberingProperties numberingProperties = Assert.IsType<NumberingProperties>(paragraphProperties.GetFirstChild<NumberingProperties>());
                Assert.Equal(2, numberingProperties.NumberingLevelReference!.Val!.Value);
                Assert.Equal(9, numberingProperties.NumberingId!.Val!.Value);

                Numbering numbering = reloaded._wordprocessingDocument!.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
                Assert.Contains(numbering.Elements<NumberingInstance>(), instance => instance.NumberID?.Value == 9);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomParagraphStyleExplicitOffRunFormattingAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string baseStyleId = "NativeDocOffBase";
            const string childStyleId = "NativeDocOffChild";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    var baseStyle = new Style { Type = StyleValues.Paragraph, StyleId = baseStyleId, CustomStyle = true };
                    baseStyle.Append(new StyleName { Val = "Native DOC Off Base" });
                    baseStyle.Append(new BasedOn { Val = WordParagraphStyles.Normal.ToStringStyle() });
                    baseStyle.Append(new StyleRunProperties(
                        new Bold(),
                        new BoldComplexScript(),
                        new Underline { Val = UnderlineValues.Single },
                        new Highlight { Val = HighlightColorValues.Yellow },
                        new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }));
                    styles.Append(baseStyle);

                    var childStyle = new Style { Type = StyleValues.Paragraph, StyleId = childStyleId, CustomStyle = true };
                    childStyle.Append(new StyleName { Val = "Native DOC Off Child" });
                    childStyle.Append(new BasedOn { Val = baseStyleId });
                    childStyle.Append(new StyleRunProperties(
                        new Bold { Val = false },
                        new BoldComplexScript { Val = false },
                        new Italic(),
                        new ItalicComplexScript(),
                        new Underline { Val = UnderlineValues.None },
                        new Highlight { Val = HighlightColorValues.None },
                        new VerticalTextAlignment { Val = VerticalPositionValues.Baseline }));
                    styles.Append(childStyle);

                    document.AddParagraph("Native custom explicit off child").SetStyleId(childStyleId);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Native custom explicit off child", paragraph.Text);
                Assert.Equal("LegacyDocNativeDOCOffChild", paragraph.StyleId);

                Styles reloadedStyles = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style baseStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCOffBase");
                StyleRunProperties baseRunProperties = Assert.IsType<StyleRunProperties>(baseStyleAfterReload.StyleRunProperties);
                Assert.NotNull(baseRunProperties.GetFirstChild<Bold>());
                Assert.NotNull(baseRunProperties.GetFirstChild<BoldComplexScript>());
                Assert.Equal(UnderlineValues.Single, baseRunProperties.GetFirstChild<Underline>()?.Val?.Value);
                Assert.Equal(HighlightColorValues.Yellow, baseRunProperties.GetFirstChild<Highlight>()?.Val?.Value);
                Assert.Equal(VerticalPositionValues.Superscript, baseRunProperties.GetFirstChild<VerticalTextAlignment>()?.Val?.Value);

                Style childStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == "LegacyDocNativeDOCOffChild");
                Assert.Equal("LegacyDocNativeDOCOffBase", childStyleAfterReload.BasedOn?.Val?.Value);
                StyleRunProperties childRunProperties = Assert.IsType<StyleRunProperties>(childStyleAfterReload.StyleRunProperties);
                Bold childBold = Assert.IsType<Bold>(childRunProperties.GetFirstChild<Bold>());
                BoldComplexScript childComplexBold = Assert.IsType<BoldComplexScript>(childRunProperties.GetFirstChild<BoldComplexScript>());
                Assert.False(childBold.Val?.Value ?? true);
                Assert.False(childComplexBold.Val?.Value ?? true);
                Assert.NotNull(childRunProperties.GetFirstChild<Italic>());
                Assert.NotNull(childRunProperties.GetFirstChild<ItalicComplexScript>());
                Assert.Equal(UnderlineValues.None, childRunProperties.GetFirstChild<Underline>()?.Val?.Value);
                Assert.Equal(HighlightColorValues.None, childRunProperties.GetFirstChild<Highlight>()?.Val?.Value);
                Assert.Equal(VerticalPositionValues.Baseline, childRunProperties.GetFirstChild<VerticalTextAlignment>()?.Val?.Value);
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInParagraphStyleNextStyleAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string heading1StyleId = WordParagraphStyles.Heading1.ToStringStyle();
            string heading2StyleId = WordParagraphStyles.Heading2.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    Style headingStyle = styles.Elements<Style>().FirstOrDefault(style => style.StyleId == heading1StyleId)
                        ?? new Style { Type = StyleValues.Paragraph, StyleId = heading1StyleId };
                    if (headingStyle.Parent == null) {
                        styles.Append(headingStyle);
                    }

                    headingStyle.GetFirstChild<NextParagraphStyle>()?.Remove();
                    headingStyle.Append(new NextParagraphStyle { Val = heading2StyleId });

                    document.AddParagraph("Native built-in next style source").SetStyle(WordParagraphStyles.Heading1);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);
                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Native built-in next style source", paragraph.Text);
                Assert.Equal(heading1StyleId, paragraph.StyleId);

                Styles reloadedStyles = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style headingStyleAfterReload = Assert.Single(reloadedStyles.Elements<Style>(), style => style.StyleId == heading1StyleId);
                Assert.Equal(heading2StyleId, headingStyleAfterReload.GetFirstChild<NextParagraphStyle>()?.Val?.Value);
            } finally {
                if (File.Exists(docPath)) {
                    File.Delete(docPath);
                }
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInStyleNumberingAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    Style headingStyle = styles.Elements<Style>().FirstOrDefault(style => style.StyleId == headingStyleId)
                        ?? new Style { Type = StyleValues.Paragraph, StyleId = headingStyleId };
                    if (headingStyle.Parent == null) {
                        styles.Append(headingStyle);
                    }

                    headingStyle.StyleParagraphProperties = new StyleParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 1 },
                            new NumberingId { Val = 7 }));

                    document.AddParagraph("Styled built-in numbered paragraph").SetStyle(WordParagraphStyles.Heading1);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.SourceFormat == WordFileFormat.Doc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Styled built-in numbered paragraph", paragraph.Text);
                Assert.Equal(headingStyleId, paragraph.StyleId);

                Styles stylesAfterReload = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style headingStyleAfterReload = Assert.Single(stylesAfterReload.Elements<Style>(), style => style.StyleId == headingStyleId);
                StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyleAfterReload.StyleParagraphProperties);
                NumberingProperties numberingProperties = Assert.IsType<NumberingProperties>(paragraphProperties.GetFirstChild<NumberingProperties>());
                Assert.Equal(1, numberingProperties.NumberingLevelReference!.Val!.Value);
                Assert.Equal(7, numberingProperties.NumberingId!.Val!.Value);

                Numbering numbering = reloaded._wordprocessingDocument!.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
                Assert.Contains(numbering.Elements<NumberingInstance>(), instance => instance.NumberID?.Value == 7);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static class LegacyDocParagraphStyleFixture {
            private const int FibLength = 0x1AA;
            private const int TextOffset = 0x200;
            private const int PapxFkpOffset = 0x400;
            private const int OleSectorSize = 512;
            private const int StyleSheetOffset = 64;
            private const int FontTableOffset = 512;
            private const ushort SprmPIstd = 0x4600;
            private const ushort SprmPFKeep = 0x2405;
            private const ushort SprmPFKeepFollow = 0x2406;
            private const ushort SprmPFPageBreakBefore = 0x2407;
            private const ushort SprmPFNoLineNumb = 0x240C;
            private const ushort SprmPFNoAutoHyph = 0x242A;
            private const ushort SprmPFKinsoku = 0x2433;
            private const ushort SprmPFWordWrap = 0x2434;
            private const ushort SprmPFOverflowPunct = 0x2435;
            private const ushort SprmPFTopLinePunct = 0x2436;
            private const ushort SprmPFAutoSpaceDE = 0x2437;
            private const ushort SprmPFAutoSpaceDN = 0x2438;
            private const ushort SprmPFBiDi = 0x2441;
            private const ushort SprmPJc = 0x2461;
            private const ushort SprmPFContextualSpacing = 0x246D;
            private const ushort SprmPFMirrorIndents = 0x2470;
            private const ushort SprmPDxaLeft = 0x840F;
            private const ushort SprmPDyaAfter = 0xA414;
            private const ushort SprmPShd80 = 0x442D;
            private const ushort SprmPFWidowControl = 0x2431;
            private const ushort SprmPWAlignFont = 0x4439;
            private const ushort SprmCFBold = 0x0835;
            private const ushort SprmCFItalic = 0x0836;
            private const ushort SprmCFStrike = 0x0837;
            private const ushort SprmCHighlight = 0x2A0C;
            private const ushort SprmCKul = 0x2A3E;
            private const ushort SprmCIss = 0x2A48;
            private const ushort SprmCIco = 0x2A42;
            private const ushort SprmCHps = 0x4A43;
            private const ushort SprmCRgFtc0 = 0x4A4F;
            private const ushort SprmCRgLid0 = 0x486D;
            private const ushort SprmCRgLid1 = 0x486E;
            private const ushort CustomStyleIndex = 10;
            private const ushort ChildStyleIndex = 11;

            internal static byte[] CreateDocWithHeadingStyles() {
                const string text = "Heading One\rHeading Two\rBody\r";
                byte[] wordDocumentStream = CreateWordDocumentStream(text);
                byte[] tableStream = CreateTableStream(text.Length);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleFontFamily() {
                const string text = "Styled Font Family\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Font Body",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(CreateCharacterSprm(SprmCRgFtc0, 0, 0)))
                });
                byte[] fontTable = CreateFontTable("Courier New");
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length,
                    FontTableOffset,
                    fontTable.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet, fontTable);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleFontFamily() {
                const string text = "Heading Font Family\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(CreateCharacterSprm(SprmCRgFtc0, 0, 0)))
                });
                byte[] fontTable = CreateFontTable("Courier New");
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length,
                    FontTableOffset,
                    fontTable.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet, fontTable);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleLanguage() {
                const string text = "Styled Language\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Language Body",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(
                            CreateCharacterSprm(SprmCRgLid0, 0x15, 0x04),
                            CreateCharacterSprm(SprmCRgLid1, 0x11, 0x04)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleLanguage() {
                const string text = "Heading Language\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(
                            CreateCharacterSprm(SprmCRgLid0, 0x15, 0x04),
                            CreateCharacterSprm(SprmCRgLid1, 0x11, 0x04)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleInheritance() {
                const string text = "Styled Custom Child\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Base",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(CreateParagraphSprm(SprmPJc, 1)),
                        characterUpx: CreateStyleCharacterUpx(CreateCharacterSprm(SprmCFBold, 1))),
                    [ChildStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Child",
                        basedOnStyleIndex: CustomStyleIndex,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(CreateCharacterSprm(SprmCFItalic, 1)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(ChildStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleNextStyle() {
                const string text = "Styled Next\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Next Source",
                        basedOnStyleIndex: 0,
                        nextStyleIndex: ChildStyleIndex,
                        paragraphUpx: null,
                        characterUpx: null),
                    [ChildStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Next Target",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(CreateCharacterSprm(SprmCFItalic, 1)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleNextStyle() {
                const string text = "Heading Next\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        nextStyleIndex: 2,
                        paragraphUpx: null,
                        characterUpx: null),
                    [2] = new LegacyDocStyleDefinition(
                        "heading 2",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStylePaginationFlags() {
                const string text = "Styled Pagination\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Pagination Body",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(
                            CreateParagraphSprm(SprmPFKeep, 1),
                            CreateParagraphSprm(SprmPFKeepFollow, 1),
                            CreateParagraphSprm(SprmPFPageBreakBefore, 1),
                            CreateParagraphSprm(SprmPFWidowControl, 1)),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStylePaginationFlags() {
                const string text = "Heading Pagination\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(
                            CreateParagraphSprm(SprmPFKeep, 1),
                            CreateParagraphSprm(SprmPFKeepFollow, 1),
                            CreateParagraphSprm(SprmPFPageBreakBefore, 1),
                            CreateParagraphSprm(SprmPFWidowControl, 1)),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleShading() {
                const string text = "Styled Shading\rBody\r";
                ushort redBackground = CreateShd80(backgroundIco: 6);
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Shaded Body",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(CreateParagraphSprm(
                            SprmPShd80,
                            (byte)(redBackground & 0xFF),
                            (byte)(redBackground >> 8))),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleShading() {
                const string text = "Heading Shading\rBody\r";
                ushort redBackground = CreateShd80(backgroundIco: 6);
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(CreateParagraphSprm(
                            SprmPShd80,
                            (byte)(redBackground & 0xFF),
                            (byte)(redBackground >> 8))),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleLayoutFlags() {
                const string text = "Styled Layout Flags\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Layout Flags Body",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(
                            CreateParagraphSprm(SprmPFNoLineNumb, 1),
                            CreateParagraphSprm(SprmPFNoAutoHyph, 1),
                            CreateParagraphSprm(SprmPFContextualSpacing, 1),
                            CreateParagraphSprm(SprmPFMirrorIndents, 1),
                            CreateParagraphSprm(SprmPFKinsoku, 1),
                            CreateParagraphSprm(SprmPFWordWrap, 1),
                            CreateParagraphSprm(SprmPFOverflowPunct, 1),
                            CreateParagraphSprm(SprmPFTopLinePunct, 1),
                            CreateParagraphSprm(SprmPFAutoSpaceDE, 1),
                            CreateParagraphSprm(SprmPFAutoSpaceDN, 1),
                            CreateParagraphSprm(SprmPFBiDi, 1),
                            CreateParagraphSprm(SprmPWAlignFont, 3, 0)),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleLayoutFlags() {
                const string text = "Heading Layout Flags\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: CreateStyleParagraphUpx(
                            CreateParagraphSprm(SprmPFNoLineNumb, 1),
                            CreateParagraphSprm(SprmPFNoAutoHyph, 1),
                            CreateParagraphSprm(SprmPFContextualSpacing, 1),
                            CreateParagraphSprm(SprmPFMirrorIndents, 1),
                            CreateParagraphSprm(SprmPFKinsoku, 1),
                            CreateParagraphSprm(SprmPFWordWrap, 1),
                            CreateParagraphSprm(SprmPFOverflowPunct, 1),
                            CreateParagraphSprm(SprmPFTopLinePunct, 1),
                            CreateParagraphSprm(SprmPFAutoSpaceDE, 1),
                            CreateParagraphSprm(SprmPFAutoSpaceDN, 1),
                            CreateParagraphSprm(SprmPFBiDi, 1),
                            CreateParagraphSprm(SprmPWAlignFont, 3, 0)),
                        characterUpx: null)
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleItalicUnderlineStrikeVerticalAndHighlight() {
                const string text = "Styled Italic Underline Strike Super Mark\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Italic Underline Strike Super Mark",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(
                            CreateCharacterSprm(SprmCFItalic, 1),
                            CreateCharacterSprm(SprmCFStrike, 1),
                            CreateCharacterSprm(SprmCIss, 1),
                            CreateCharacterSprm(SprmCHighlight, 7),
                            CreateCharacterSprm(SprmCKul, 1)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithBuiltInParagraphStyleItalicUnderlineStrikeVerticalAndHighlight() {
                const string text = "Heading Italic Underline Strike Super Mark\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [1] = new LegacyDocStyleDefinition(
                        "heading 1",
                        basedOnStyleIndex: 0,
                        paragraphUpx: null,
                        characterUpx: CreateStyleCharacterUpx(
                            CreateCharacterSprm(SprmCFItalic, 1),
                            CreateCharacterSprm(SprmCFStrike, 1),
                            CreateCharacterSprm(SprmCIss, 1),
                            CreateCharacterSprm(SprmCHighlight, 7),
                            CreateCharacterSprm(SprmCKul, 1)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyle() {
                const string text = "Styled Custom\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, string> {
                    [CustomStyleIndex] = "Custom Body"
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            internal static byte[] CreateDocWithCustomParagraphStyleFormatting() {
                const string text = "Styled Custom Formatting\rBody\r";
                byte[] styleSheet = CreateStyleSheet(new Dictionary<ushort, LegacyDocStyleDefinition> {
                    [CustomStyleIndex] = new LegacyDocStyleDefinition(
                        "Custom Formatted Body",
                        basedOnStyleIndex: 1,
                        paragraphUpx: CreateStyleParagraphUpx(
                            CreateParagraphSprm(SprmPJc, 1),
                            CreateParagraphSprm(SprmPDyaAfter, 0xF0, 0x00),
                            CreateParagraphSprm(SprmPDxaLeft, 0xD0, 0x02)),
                        characterUpx: CreateStyleCharacterUpx(
                            CreateCharacterSprm(SprmCFBold, 1),
                            CreateCharacterSprm(SprmCHps, 28, 0),
                            CreateCharacterSprm(SprmCIco, 6)))
                });
                byte[] wordDocumentStream = CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(CustomStyleIndex)
                    },
                    styleSheet.Length);
                byte[] tableStream = CreateTableStream(text.Length, styleSheet);

                using var package = new MemoryStream();
                using (RootStorage root = RootStorage.Create(package, Version.V3, StorageModeFlags.LeaveOpen)) {
                    WriteStream(root, "WordDocument", wordDocumentStream);
                    WriteStream(root, "1Table", tableStream);
                }

                return package.ToArray();
            }

            private static byte[] CreateWordDocumentStream(string text) {
                return CreateWordDocumentStream(
                    text,
                    new Dictionary<int, byte[]> {
                        [0] = CreateParagraphStylePapx(1),
                        [1] = CreateParagraphStyleSprmPapx(2)
                    },
                    styleSheetLength: 0);
            }

            private static byte[] CreateWordDocumentStream(
                string text,
                IReadOnlyDictionary<int, byte[]> papxByParagraphIndex,
                int styleSheetLength,
                int fontTableOffset = 0,
                int fontTableLength = 0) {
                byte[] textBytes = System.Text.Encoding.Unicode.GetBytes(text);
                var stream = new byte[Math.Max(PapxFkpOffset + OleSectorSize, TextOffset + textBytes.Length)];
                WriteUInt16(stream, 0x00, 0xA5EC);
                WriteUInt16(stream, 0x02, 0x00D9);
                WriteUInt16(stream, 0x0A, 0x0200);
                WriteInt32(stream, 0x4C, text.Length);
                if (styleSheetLength > 0) {
                    WriteInt32(stream, 0xA2, StyleSheetOffset);
                    WriteInt32(stream, 0xA6, styleSheetLength);
                }

                WriteInt32(stream, 0x102, 21);
                WriteInt32(stream, 0x106, 12);
                if (fontTableLength > 0) {
                    WriteInt32(stream, 0x112, fontTableOffset);
                    WriteInt32(stream, 0x116, fontTableLength);
                }

                WriteInt32(stream, 0x1A2, 0);
                WriteInt32(stream, 0x1A6, 21);
                Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);

                WritePapxFkp(stream, CreateParagraphPositions(text), papxByParagraphIndex);

                if (stream.Length < FibLength) {
                    Array.Resize(ref stream, FibLength);
                }

                return stream;
            }

            private static byte[] CreateTableStream(int characterCount, byte[]? styleSheet = null, byte[]? fontTable = null) {
                int length = Math.Max(
                    styleSheet == null ? 33 : StyleSheetOffset + styleSheet.Length,
                    fontTable == null ? 33 : FontTableOffset + fontTable.Length);
                var table = new byte[length];
                table[0] = 0x02;
                WriteInt32(table, 1, 16);
                WriteInt32(table, 5, 0);
                WriteInt32(table, 9, characterCount);
                WriteUInt16(table, 13, 0);
                WriteUInt32(table, 15, TextOffset);
                WriteUInt16(table, 19, 0);

                int papxPlcOffset = 21;
                WriteInt32(table, papxPlcOffset, TextOffset);
                WriteInt32(table, papxPlcOffset + 4, TextOffset + (characterCount * 2));
                WriteInt32(table, papxPlcOffset + 8, PapxFkpOffset / OleSectorSize);
                if (styleSheet != null) {
                    Buffer.BlockCopy(styleSheet, 0, table, StyleSheetOffset, styleSheet.Length);
                }

                if (fontTable != null) {
                    Buffer.BlockCopy(fontTable, 0, table, FontTableOffset, fontTable.Length);
                }

                return table;
            }

            private static int[] CreateParagraphPositions(string text) {
                var positions = new List<int> { TextOffset };
                int characterOffset = 0;
                foreach (char character in text) {
                    characterOffset++;
                    if (character == '\r') {
                        positions.Add(TextOffset + (characterOffset * 2));
                    }
                }

                return positions.ToArray();
            }

            private static byte[] CreateStyleSheet(IReadOnlyDictionary<ushort, string> styleNamesByIndex) {
                return CreateStyleSheet(styleNamesByIndex.ToDictionary(
                    pair => pair.Key,
                    pair => new LegacyDocStyleDefinition(pair.Value, basedOnStyleIndex: 0, paragraphUpx: null, characterUpx: null)));
            }

            private static byte[] CreateStyleSheet(IReadOnlyDictionary<ushort, LegacyDocStyleDefinition> stylesByIndex) {
                ushort cstd = checked((ushort)(stylesByIndex.Keys.Max() + 1));
                var bytes = new List<byte>();
                WriteUInt16(bytes, 18);
                WriteUInt16(bytes, cstd);
                WriteUInt16(bytes, 10);
                for (int i = 0; i < 7; i++) {
                    WriteUInt16(bytes, 0);
                }

                for (ushort index = 0; index < cstd; index++) {
                    if (!stylesByIndex.TryGetValue(index, out LegacyDocStyleDefinition? definition)) {
                        WriteUInt16(bytes, 0);
                        continue;
                    }

                    byte[] std = CreateParagraphStyleDefinition(definition);
                    WriteUInt16(bytes, checked((ushort)std.Length));
                    bytes.AddRange(std);
                    if (bytes.Count % 2 != 0) {
                        bytes.Add(0);
                    }
                }

                return bytes.ToArray();
            }

            private static byte[] CreateParagraphStyleDefinition(LegacyDocStyleDefinition definition) {
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(definition.Name);
                var bytes = new List<byte>();
                WriteUInt16(bytes, 0x0FFE);
                WriteUInt16(bytes, 0x0001 | (definition.BasedOnStyleIndex << 4));
                int upxCount = (definition.ParagraphUpx == null && definition.CharacterUpx == null) ? 0 : 2;
                WriteUInt16(bytes, upxCount | (definition.NextStyleIndex << 4));
                WriteUInt16(bytes, 0);
                WriteUInt16(bytes, 0);
                WriteUInt16(bytes, checked((ushort)definition.Name.Length));
                bytes.AddRange(nameBytes);
                WriteUInt16(bytes, 0);
                if (definition.ParagraphUpx != null || definition.CharacterUpx != null) {
                    WriteLengthPrefixedUpx(bytes, definition.ParagraphUpx ?? Array.Empty<byte>());
                    WriteLengthPrefixedUpx(bytes, definition.CharacterUpx ?? Array.Empty<byte>());
                }

                return bytes.ToArray();
            }

            private static void WritePapxFkp(byte[] stream, int[] fileParagraphPositions, IReadOnlyDictionary<int, byte[]> papxByParagraphIndex) {
                const int bxLength = 13;
                int paragraphCount = fileParagraphPositions.Length - 1;
                for (int i = 0; i < fileParagraphPositions.Length; i++) {
                    WriteInt32(stream, PapxFkpOffset + (i * 4), fileParagraphPositions[i]);
                }

                int rgbxOffset = PapxFkpOffset + (fileParagraphPositions.Length * 4);
                int papxOffset = 0x180;
                for (int i = 0; i < paragraphCount; i++) {
                    if (!papxByParagraphIndex.TryGetValue(i, out byte[]? papx)) {
                        continue;
                    }

                    papxOffset = AlignToEven(papxOffset);
                    stream[rgbxOffset + (i * bxLength)] = checked((byte)(papxOffset / 2));
                    Buffer.BlockCopy(papx, 0, stream, PapxFkpOffset + papxOffset, papx.Length);
                    papxOffset += papx.Length;
                }

                stream[PapxFkpOffset + OleSectorSize - 1] = checked((byte)paragraphCount);
            }

            private static byte[] CreateParagraphStylePapx(ushort styleIndex) {
                return CreateParagraphPropertiesPapx(styleIndex);
            }

            private static byte[] CreateParagraphStyleSprmPapx(ushort styleIndex) {
                return CreateParagraphPropertiesPapx(0, CreateParagraphSprm(SprmPIstd, (byte)(styleIndex & 0xFF), (byte)(styleIndex >> 8)));
            }

            private static byte[] CreateParagraphPropertiesPapx(ushort baseStyleIndex, params byte[][] sprms) {
                var grpprl = new List<byte> {
                    (byte)(baseStyleIndex & 0xFF),
                    (byte)(baseStyleIndex >> 8)
                };

                foreach (byte[] sprm in sprms) {
                    grpprl.AddRange(sprm);
                }

                if (grpprl.Count % 2 != 0) {
                    grpprl.Add(0);
                }

                var papx = new byte[grpprl.Count + 2];
                papx[0] = 0;
                papx[1] = checked((byte)(grpprl.Count / 2));
                grpprl.CopyTo(papx, 2);
                return papx;
            }

            private static byte[] CreateParagraphSprm(ushort sprm, params byte[] operand) {
                var bytes = new byte[2 + operand.Length];
                WriteUInt16(bytes, 0, sprm);
                Buffer.BlockCopy(operand, 0, bytes, 2, operand.Length);
                return bytes;
            }

            private static ushort CreateShd80(byte backgroundIco) {
                return (ushort)(backgroundIco << 5);
            }

            private static byte[] CreateStyleParagraphUpx(params byte[][] sprms) {
                var bytes = new List<byte>();
                foreach (byte[] sprm in sprms) {
                    bytes.AddRange(sprm);
                }

                return bytes.ToArray();
            }

            private static byte[] CreateStyleCharacterUpx(params byte[][] sprms) {
                var bytes = new List<byte>();
                foreach (byte[] sprm in sprms) {
                    bytes.AddRange(sprm);
                }

                return bytes.ToArray();
            }

            private static byte[] CreateCharacterSprm(ushort sprm, params byte[] operand) {
                var bytes = new byte[2 + operand.Length];
                WriteUInt16(bytes, 0, sprm);
                Buffer.BlockCopy(operand, 0, bytes, 2, operand.Length);
                return bytes;
            }

            private static byte[] CreateFontTable(params string[] fontFamilies) {
                var bytes = new List<byte>();
                WriteUInt16(bytes, checked((ushort)fontFamilies.Length));
                WriteUInt16(bytes, 0);
                foreach (string fontFamily in fontFamilies) {
                    byte[] ffn = CreateFfn(fontFamily);
                    bytes.Add(checked((byte)ffn.Length));
                    bytes.AddRange(ffn);
                }

                return bytes.ToArray();
            }

            private static byte[] CreateFfn(string fontFamily) {
                byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(fontFamily + '\0');
                var ffn = new byte[39 + nameBytes.Length];
                ffn[1] = 0x90;
                ffn[2] = 0x01;
                Buffer.BlockCopy(nameBytes, 0, ffn, 39, nameBytes.Length);
                return ffn;
            }

            private static void WriteLengthPrefixedUpx(List<byte> bytes, byte[] upx) {
                WriteUInt16(bytes, checked((ushort)upx.Length));
                bytes.AddRange(upx);
                if (bytes.Count % 2 != 0) {
                    bytes.Add(0);
                }
            }

            private static void WriteStream(RootStorage root, string name, byte[] bytes) {
                using CfbStream stream = root.CreateStream(name);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static int AlignToEven(int value) {
                return value % 2 == 0 ? value : value + 1;
            }

            private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
                bytes[offset] = (byte)(value & 0xFF);
                bytes[offset + 1] = (byte)(value >> 8);
            }

            private static void WriteUInt16(List<byte> bytes, int value) {
                bytes.Add((byte)(value & 0xFF));
                bytes.Add((byte)(value >> 8));
            }

            private static void WriteInt32(byte[] bytes, int offset, int value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private static void WriteUInt32(byte[] bytes, int offset, uint value) {
                bytes[offset] = (byte)value;
                bytes[offset + 1] = (byte)(value >> 8);
                bytes[offset + 2] = (byte)(value >> 16);
                bytes[offset + 3] = (byte)(value >> 24);
            }

            private sealed class LegacyDocStyleDefinition {
                internal LegacyDocStyleDefinition(string name, ushort basedOnStyleIndex, byte[]? paragraphUpx, byte[]? characterUpx) {
                    Name = name;
                    BasedOnStyleIndex = basedOnStyleIndex;
                    NextStyleIndex = 0;
                    ParagraphUpx = paragraphUpx;
                    CharacterUpx = characterUpx;
                }

                internal LegacyDocStyleDefinition(string name, ushort basedOnStyleIndex, ushort nextStyleIndex, byte[]? paragraphUpx, byte[]? characterUpx) {
                    Name = name;
                    BasedOnStyleIndex = basedOnStyleIndex;
                    NextStyleIndex = nextStyleIndex;
                    ParagraphUpx = paragraphUpx;
                    CharacterUpx = characterUpx;
                }

                internal string Name { get; }

                internal ushort BasedOnStyleIndex { get; }

                internal ushort NextStyleIndex { get; }

                internal byte[]? ParagraphUpx { get; }

                internal byte[]? CharacterUpx { get; }
            }
        }
    }
}
