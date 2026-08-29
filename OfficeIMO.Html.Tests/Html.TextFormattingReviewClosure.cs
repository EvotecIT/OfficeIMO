using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class HtmlTextFormattingReviewClosureTests {
    [Fact]
    public void PowerPointSemanticHtmlResolvesInheritedTypographyAndDirectOffOverrides() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointParagraph paragraph = source.AddSlide().AddTextBox("Inherited")
            .Paragraphs.Single();
        A.ParagraphProperties paragraphProperties = paragraph.Paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        var defaults = new A.DefaultRunProperties {
            Bold = true,
            Italic = true,
            Underline = A.TextUnderlineValues.Wavy,
            Strike = A.TextStrikeValues.DoubleStrike,
            Capital = A.TextCapsValues.Small,
            Baseline = 30000,
            FontSize = 1800,
            Language = "pl-PL"
        };
        defaults.Append(new A.LatinFont { Typeface = "Aptos Display" });
        defaults.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "336699" }));
        paragraphProperties.Append(defaults);

        PowerPointTextRun directOff = paragraph.AddRun(" Plain");
        directOff.Run.RunProperties = new A.RunProperties {
            Bold = false,
            Italic = false,
            Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike,
            Capital = A.TextCapsValues.None,
            Baseline = 0
        };

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());

        Assert.Contains("font-weight:700", html, StringComparison.Ordinal);
        Assert.Contains("font-style:italic", html, StringComparison.Ordinal);
        Assert.Contains("font-family:&#39;Aptos Display&#39;", html, StringComparison.Ordinal);
        Assert.Contains("font-size:18pt", html, StringComparison.Ordinal);
        Assert.Contains("color:#336699", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-underline=\"Wavy\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-strike=\"Double\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-capitalization=\"SmallCaps\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-baseline-percent=\"30\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-language=\"pl-PL\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-underline=\"None\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-strike=\"None\"", html, StringComparison.Ordinal);

        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();
        PowerPointTextRun[] runs = Assert.Single(Assert.Single(Assert.Single(imported.Slides).TextBoxes).Paragraphs)
            .Runs.ToArray();
        Assert.Equal(2, runs.Length);
        Assert.True(runs[0].Bold);
        Assert.True(runs[0].Italic);
        Assert.Equal(PowerPointUnderlineStyle.Wavy, runs[0].UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, runs[0].StrikeStyle);
        Assert.Equal(PowerPointCapitalization.SmallCaps, runs[0].Capitalization);
        Assert.Equal(30D, runs[0].BaselinePercent);
        Assert.Equal(18D, runs[0].FontSizePoints);
        Assert.Equal("Aptos Display", runs[0].FontName);
        Assert.Equal("336699", runs[0].Color);
        Assert.Equal("pl-PL", runs[0].Language);
        Assert.False(runs[1].Bold);
        Assert.False(runs[1].Italic);
        Assert.Equal(PowerPointUnderlineStyle.None, runs[1].UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.None, runs[1].StrikeStyle);
        Assert.Equal(PowerPointCapitalization.None, runs[1].Capitalization);
        Assert.Equal(0D, runs[1].BaselinePercent);
    }

    [Fact]
    public void PowerPointSemanticHtmlResolvesMappedThemeColors() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointSlide slide = source.AddSlide();
        PowerPointParagraph paragraph = slide.AddTextBox("Theme text").Paragraphs[0];
        PowerPointTextRun run = paragraph.Runs[0];
        run.Run.RunProperties = new A.RunProperties();
        run.RunProperties.Append(new A.SolidFill(new A.SchemeColor {
            Val = A.SchemeColorValues.Text1
        }));
        PowerPointTextRun systemRun = paragraph.AddRun(" System text");
        systemRun.Run.RunProperties = new A.RunProperties();
        systemRun.RunProperties.Append(new A.SolidFill(new A.SystemColor {
            Val = A.SystemColorValues.WindowText,
            LastColor = "884422"
        }));

        var master = slide.SlidePart.SlideLayoutPart!.SlideMasterPart!;
        master.SlideMaster!.ColorMap!.Text1 = A.ColorSchemeIndexValues.Accent2;
        A.Accent2Color accent = master.ThemePart!.Theme!.ThemeElements!.ColorScheme!
            .GetFirstChild<A.Accent2Color>()!;
        accent.RemoveAllChildren();
        accent.Append(new A.RgbColorModelHex { Val = "336699" });

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());

        Assert.Contains("color:#336699", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("color:#884422", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PowerPointSemanticHtmlResolvesThemeFontTokensToRealFamilies() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        source.SetThemeLatinFonts("Theme Major", "Theme Minor");
        PowerPointTextRun run = source.AddSlide().AddTextBox("Theme font").Paragraphs[0].Runs[0];
        run.Run.RunProperties = new A.RunProperties();
        run.RunProperties.Append(new A.LatinFont { Typeface = "+mn-lt" });

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());

        Assert.Contains("font-family:&#39;Theme Minor&#39;", html, StringComparison.Ordinal);
        Assert.DoesNotContain("+mn-lt", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PowerPointSemanticHtmlRoundTripPreservesParagraphBreaksAndRunFormatting() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointTextBox textBox = source.AddSlide().AddTextBox("First");
        textBox.Paragraphs[0].Runs[0].Bold = true;
        PowerPointParagraph second = textBox.AddParagraph("Second");
        second.Runs[0].Italic = true;

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());
        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();

        PowerPointTextBox actual = Assert.Single(Assert.Single(imported.Slides).TextBoxes);
        Assert.Equal(2, actual.Paragraphs.Count);
        Assert.Equal("First", actual.Paragraphs[0].Text);
        Assert.Equal("Second", actual.Paragraphs[1].Text);
        Assert.True(actual.Paragraphs[0].Runs[0].Bold);
        Assert.True(actual.Paragraphs[1].Runs[0].Italic);
    }

    [Fact]
    public void PowerPointSemanticHtmlIncludesExplicitLineBreaksAndFieldsInOrder() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointParagraph paragraph = source.AddSlide().AddTextBox("Before")
            .Paragraphs.Single();
        paragraph.AddLineBreak();
        paragraph.AddField("27 August 2026", "datetime1", "{11111111-1111-1111-1111-111111111111}");

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());

        Assert.Contains("Before</span><br data-officeimo-powerpoint-inline-break=\"true\"><span data-officeimo-powerpoint-field=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-field-id=\"{11111111-1111-1111-1111-111111111111}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-field-type=\"datetime1\"", html, StringComparison.Ordinal);
        Assert.Contains(">27 August 2026</span>", html, StringComparison.Ordinal);

        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();
        PowerPointParagraph actual = Assert.Single(Assert.Single(Assert.Single(imported.Slides).TextBoxes).Paragraphs);
        Assert.Collection(actual.InlineNodes,
            node => Assert.Equal(PowerPointParagraphInlineKind.Run, node.Kind),
            node => Assert.Equal(PowerPointParagraphInlineKind.LineBreak, node.Kind),
            node => Assert.Equal(PowerPointParagraphInlineKind.Field, node.Kind));
    }

    [Fact]
    public void PowerPointSemanticHtmlKeepsUnderlineStrikeAndEscapedFontFamilyIndependent() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointTextRun run = source.AddSlide().AddTextBox("Styled").Paragraphs.Single().Runs.Single();
        run.UnderlineStyle = PowerPointUnderlineStyle.Wavy;
        run.StrikeStyle = PowerPointStrikeStyle.Double;
        run.FontName = "O'Brien Sans";

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());

        Assert.Contains("data-officeimo-powerpoint-font-family=\"O&#39;Brien Sans\"", html, StringComparison.Ordinal);
        Assert.Contains("text-decoration-line:underline;text-decoration-style:wavy", html, StringComparison.Ordinal);
        Assert.Contains("text-decoration-line:line-through;text-decoration-style:double", html, StringComparison.Ordinal);

        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();
        PowerPointTextRun actual = Assert.Single(Assert.Single(Assert.Single(imported.Slides).TextBoxes).Paragraphs).Runs.Single();
        Assert.Equal(PowerPointUnderlineStyle.Wavy, actual.UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, actual.StrikeStyle);
        Assert.Equal("O'Brien Sans", actual.FontName);
    }

    [Fact]
    public void PowerPointSemanticHtmlFallsBackForMixedAndNestedOrdinarySpanContent() {
        const string html = "<section class='officeimo-slide'><p><span><strong>Bold</strong></span> tail</p></section>";

        using PowerPointPresentation imported = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult()
            .RequireValue();

        PowerPointTextBox textBox = Assert.Single(Assert.Single(imported.Slides).TextBoxes);
        Assert.Equal("Bold tail", textBox.Text);
        Assert.True(textBox.Paragraphs.Single().Runs.First().Bold);
    }

    [Fact]
    public void WordStyleDefinitionsRenderUnderlineAndDoubleStrikeWithIndependentPatterns() {
        using WordDocument source = WordDocument.Create();
        var style = new Style { Type = StyleValues.Character, StyleId = "Decorated" };
        style.Append(new StyleName { Val = "Decorated" });
        var properties = new StyleRunProperties();
        properties.Append(new Underline { Val = UnderlineValues.Wave });
        properties.Append(new DoubleStrike());
        style.Append(properties);
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        source.AddParagraph().AddText("Styled").SetCharacterStyleId("Decorated");

        string html = source.ToHtml(new WordToHtmlOptions { IncludeRunClasses = true });

        Assert.Contains("class=\"Decorated\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-word-underline=\"Wave\"", html, StringComparison.Ordinal);
        Assert.Contains("text-decoration-line:underline;text-decoration-style:wavy", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-word-double-strike=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("text-decoration-line:line-through;text-decoration-style:double", html, StringComparison.Ordinal);
        string styleRule = Assert.Single(html.Split('\n'), line => line.Contains(".Decorated {", StringComparison.Ordinal));
        Assert.DoesNotContain("text-decoration", styleRule, StringComparison.Ordinal);

        using WordDocument imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToWordDocumentResult()
            .RequireValue();
        WordParagraph actual = Assert.Single(imported.Paragraphs);
        Assert.Equal(WordUnderlineStyle.Wave, actual.Underline);
        Assert.True(actual.DoubleStrike);
    }

    [Fact]
    public void WordParagraphStyleDecorationsAndScriptAllowDirectRunResets() {
        using WordDocument source = WordDocument.Create();
        var style = new Style { Type = StyleValues.Paragraph, StyleId = "ParagraphDecoratedScript" };
        style.Append(new StyleName { Val = "Paragraph Decorated Script" });
        var properties = new StyleRunProperties();
        properties.Append(new Underline { Val = UnderlineValues.Single });
        properties.Append(new Strike());
        properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        style.Append(properties);
        source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);

        WordParagraph paragraph = source.AddParagraph();
        paragraph.SetStyleId("ParagraphDecoratedScript");
        paragraph.AddText("Inherited");
        WordParagraph reset = paragraph.AddText("Reset")
            .SetUnderline(WordUnderlineStyle.None)
            .SetVerticalTextAlignment(WordVerticalTextPosition.Baseline);
        reset._runProperties!.Strike = new Strike { Val = false };

        string html = source.ToHtml(new WordToHtmlOptions { IncludeParagraphClasses = true });
        IElement output = HtmlDocumentParser.ParseDocument(html)
            .QuerySelector("p.ParagraphDecoratedScript")!;

        Assert.Equal("Inherited", Assert.Single(output.QuerySelectorAll("sup")).TextContent);
        Assert.Equal("Inherited", Assert.Single(output.QuerySelectorAll("u")).TextContent);
        IHtmlCollection<IElement> strikeElements = output.QuerySelectorAll("s");
        Assert.NotEmpty(strikeElements);
        Assert.All(strikeElements, element => Assert.Equal("Inherited", element.TextContent));
        Assert.Equal("InheritedReset", output.TextContent);
        string styleRule = Assert.Single(html.Split('\n'), line => line.Contains(".ParagraphDecoratedScript {", StringComparison.Ordinal));
        Assert.DoesNotContain("text-decoration", styleRule, StringComparison.Ordinal);
        Assert.DoesNotContain("vertical-align", styleRule, StringComparison.Ordinal);
    }

    [Fact]
    public void LegacyPowerPointSemanticHeadingIsNotPairedWithATextBox() {
        const string html = """
            <section class="officeimo-slide">
              <h2>Slide 1</h2>
              <p>Actual first box</p>
              <p>Actual second box</p>
            </section>
            """;

        using PowerPointPresentation imported = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult()
            .RequireValue();

        string[] texts = Assert.Single(imported.Slides).TextBoxes.Select(box => box.Text).ToArray();
        Assert.Equal(new[] { "Actual first box", "Actual second box" }, texts);
    }

    [Fact]
    public void OneNoteSemanticHtmlRendersInlineMathFromTheExpression() {
        OneNoteSection section = CreateOneNoteSection(out OneNoteParagraph paragraph);
        paragraph.AddMath(OfficeIMO.Drawing.OfficeMath.Fraction(
            OfficeIMO.Drawing.OfficeMath.Identifier("x"), OfficeIMO.Drawing.OfficeMath.Number("2")));

        string html = section.ToHtmlDocument();

        Assert.Contains("class=\"officeimo-onenote-math\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-math-format=\"latex\"", html, StringComparison.Ordinal);
        Assert.Contains("\\frac{x}{2}", html, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteSemanticHtmlDoesNotResolveMissingBinaryPayloads() {
        OneNoteSection section = CreateOneNoteSection(out _);
        OneNotePage page = Assert.Single(section.Pages);
        page.DirectContent.Add(new OneNoteImage { FileName = "missing.png", AltText = "Missing" });
        int resolverCalls = 0;

        string html = section.ToHtmlDocument(new OfficeIMO.OneNote.Markdown.OneNoteMarkdownOptions {
            AssetUriResolver = element => {
                resolverCalls++;
                return element.Payload!.Length.ToString();
            }
        });

        Assert.Equal(0, resolverCalls);
        Assert.Contains("officeimo-onenote-image-placeholder", html, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteSemanticHtmlReportsAssetUrisRejectedByPolicy() {
        OneNoteSection section = CreateOneNoteSection(out _);
        OneNotePage page = Assert.Single(section.Pages);
        page.DirectContent.Add(new OneNoteImage {
            FileName = "blocked.png",
            AltText = "Blocked",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });

        HtmlTextConversionResult result = section.ToHtmlDocumentResult(new OfficeIMO.OneNote.Markdown.OneNoteMarkdownOptions {
            AssetUriResolver = _ => "javascript:alert(1)"
        });

        Assert.Contains("officeimo-onenote-image-placeholder", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("javascript:", result.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "ONENOTE_HTML_ASSET_URI_REJECTED_BY_POLICY"
            && diagnostic.Source == "blocked.png"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);

        HtmlTextConversionResult allowed = section.ToHtmlDocumentResult(new OfficeIMO.OneNote.Markdown.OneNoteMarkdownOptions {
            AssetUriResolver = _ => "https://example.com/allowed.png"
        });
        Assert.Contains("src=\"https://example.com/allowed.png\"", allowed.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(allowed.Report.Diagnostics,
            diagnostic => diagnostic.Code == "ONENOTE_HTML_ASSET_URI_REJECTED_BY_POLICY");
    }

    [Fact]
    public void OneNoteHtmlDiagnosticsSuppressPreservedStylesButRetainSpacingLoss() {
        OneNoteSection section = CreateOneNoteSection(out OneNoteParagraph paragraph);
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Styled" });
        paragraph.Runs[0].Style.FontFamily = "Aptos";
        paragraph.Runs[0].Style.Underline = true;

        HtmlTextConversionResult styled = section.ToHtmlDocumentResult();
        Assert.DoesNotContain(styled.Report.Diagnostics,
            diagnostic => diagnostic.Code is "ONENOTE_MARKDOWN_FORMATTING_SIMPLIFIED" or "ONENOTE_HTML_FORMATTING_SIMPLIFIED");

        paragraph.Style.SpaceBefore = 12D;
        HtmlTextConversionResult spaced = section.ToHtmlDocumentResult();
        Assert.Contains(spaced.Report.Diagnostics,
            diagnostic => diagnostic.Code == "ONENOTE_HTML_FORMATTING_SIMPLIFIED"
                && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    private static OneNoteSection CreateOneNoteSection(out OneNoteParagraph paragraph) {
        var section = new OneNoteSection { Name = "Review" };
        var page = new OneNotePage { Title = "Review" };
        paragraph = new OneNoteParagraph();
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }
}
