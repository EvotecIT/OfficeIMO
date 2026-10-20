using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("absolute", "div")]
    [InlineData("absolute", "span")]
    [InlineData("fixed", "div")]
    [InlineData("fixed", "span")]
    public void HtmlPositionedText_RetainsSourceOrderAmongOrdinarySiblings(string position, string tag) {
        string html = PositionedLogicalStyles
            + PositionedLogicalElement(position, tag, "Early", 80)
            + "<p>Before</p>"
            + PositionedLogicalElement(position, tag, "Middle", 100)
            + "<p>After</p>"
            + PositionedLogicalElement(position, tag, "Late", 120);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("EarlyBeforeMiddleAfterLate", WithoutPositionedLogicalWhitespace(rendered.Text));
        HtmlRenderText[] paint = EnumerateRenderVisuals(Assert.Single(rendered.Pages).Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText before = Assert.Single(paint, text => text.Text == "Before");
        HtmlRenderText after = Assert.Single(paint, text => text.Text == "After");
        foreach (var item in new[] { ("Early", 80D), ("Middle", 100D), ("Late", 120D) }) {
            HtmlRenderText positioned = Assert.Single(paint, text => text.Text == item.Item1);
            Assert.Equal(20D, positioned.X, 3);
            Assert.Equal(item.Item2, positioned.Y - before.Y, 3);
            Assert.True(Array.IndexOf(paint, after) < Array.IndexOf(paint, positioned));
        }
        Assert.Equal(20D, after.Y - before.Y, 3);
    }

    [Theory]
    [InlineData("absolute", "", HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData("absolute", "opacity:.75", HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData("absolute", "overflow:hidden", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("absolute", "clip-path:inset(0)", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("fixed", "", HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData("fixed", "opacity:.75", HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData("fixed", "overflow:hidden", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("fixed", "clip-path:inset(0)", HtmlRenderIntentProfile.PrintPaged)]
    public void HtmlPositionedText_RetainsSourceOrderInsideContainingGroup(
        string position, string effect, HtmlRenderIntentProfile profile) {
        // The transform supplies the fixed containing block, so both positioned
        // modes exercise the same nested semantic, effect and clipping owners.
        string html = PositionedLogicalStyles + "<p>Lead</p>"
            + "<section style='position:relative;transform:translateX(0);width:240px;height:160px;" + effect + "'>"
            + PositionedLogicalElement(position, "div", "Early", 80)
            + "<p>Before</p>"
            + PositionedLogicalElement(position, "div", "Middle", 100)
            + "<p>After</p>"
            + PositionedLogicalElement(position, "div", "Late", 120)
            + "</section><p>Tail</p>";
        HtmlRenderResult result = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, options: PositionedLogicalOptions()));

        Assert.Single(result.Document.Pages);
        Assert.Equal("LeadEarlyBeforeMiddleAfterLateTail", WithoutPositionedLogicalWhitespace(result.Document.Text));
    }

    [Theory]
    [InlineData("absolute", "section", "")]
    [InlineData("fixed", "section", "")]
    [InlineData("absolute", "div", "opacity:.8;overflow:hidden")]
    [InlineData("fixed", "div", "opacity:.8;overflow:hidden")]
    public void HtmlPositionedText_LiftedDescendantInterleavesStructuralEnvelopeText(
        string position, string tag, string effect) {
        string html = PositionedLogicalStyles + "<p>Lead</p><" + tag + " style='height:160px;" + effect + "'>"
            + "<p>Before</p>" + PositionedLogicalElement(position, "span", "Positioned", 80)
            + "<p>After</p></" + tag + "><p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("LeadBeforePositionedAfterTail", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPositionedText_RetainsOrderBesideDrawingBackedMath(string position) {
        string html = PositionedLogicalStyles + "<p>Before</p><math><mi>x</mi></math>"
            + PositionedLogicalElement(position, "span", "Positioned", 80) + "<p>After</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("BeforexPositionedAfter", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPositionedText_RetainsOrderInsideBidirectionalParagraph(string position) {
        string html = PositionedLogicalStyles + "<p dir='rtl'>שלום Before "
            + "<span style='position:" + position + ";top:80px'>Positioned</span> After עולם</p><p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, Margins = HtmlRenderMargins.All(0D) });

        Assert.Equal("שלוםBeforePositionedAfterעולםTail", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute", false)]
    [InlineData("fixed", false)]
    [InlineData("absolute", true)]
    [InlineData("fixed", true)]
    public void HtmlPositionedText_RetainsOrderBesideMissingImageAlternative(string position, bool longAlternative) {
        string alternative = longAlternative ? "Start" + new string('a', 1100) + "End" : "Alt";
        string html = PositionedLogicalStyles + "<p>Lead</p><img src='missing.png' alt='" + alternative
            + "' style='width:40px;height:20px'>" + PositionedLogicalElement(position, "span", "Positioned", 80)
            + "<p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("Lead" + alternative + "PositionedTail", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute", "outside", false)]
    [InlineData("fixed", "outside", false)]
    [InlineData("absolute", "inside", false)]
    [InlineData("fixed", "inside", false)]
    [InlineData("absolute", "outside", true)]
    [InlineData("fixed", "outside", true)]
    public void HtmlPositionedText_RetainsOrderBesideDefaultListMarker(
        string position, string markerPosition, bool blockChild) {
        string item = blockChild ? "<p>List</p>" : "List";
        string html = PositionedLogicalStyles + "<p>Lead</p><ul style='list-style-position:" + markerPosition
            + "'><li>" + item + "</li></ul>" + PositionedLogicalElement(position, "span", "Positioned", 80)
            + "<p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("Lead•ListPositionedTail", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPositionedText_RetainsOrderBesidePagedFootnote(string position) {
        string html = PositionedLogicalStyles + "<style>.note{float:footnote}</style>"
            + "<p>Lead<span class='note'>Note</span></p>"
            + PositionedLogicalElement(position, "span", "Positioned", 80) + "<p>Tail</p>";
        HtmlRenderOptions options = PositionedLogicalOptions();
        options.Mode = HtmlRenderMode.Paged;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Single(rendered.Pages);
        Assert.Equal("Lead11NotePositionedTail", WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute", false)]
    [InlineData("fixed", false)]
    [InlineData("absolute", true)]
    [InlineData("fixed", true)]
    public void HtmlPositionedText_AuthoredFootnoteCallPrecedesPlainAndMixedMarker(string position, bool mixedMarker) {
        string marker = "'M'" + (mixedMarker ? " url('data:image/png;base64," + ZeroFontFootnotePixel + "')" : string.Empty);
        string html = PositionedLogicalStyles + "<style>.note{float:footnote}"
            + ".note::footnote-call{content:'C'}.note::footnote-marker{content:" + marker + "}</style>"
            + "<p>Lead<span class='note'>Note</span></p>"
            + PositionedLogicalElement(position, "span", "Positioned", 80) + "<p>Tail</p>";
        HtmlRenderOptions options = PositionedLogicalOptions();
        options.Mode = HtmlRenderMode.Paged;
        options.ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Single(rendered.Pages);
        Assert.Equal("LeadCMNotePositionedTail", WithoutPositionedLogicalWhitespace(rendered.Text));
        Assert.Equal(mixedMarker ? 1 : 0,
            rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderImage>().Count());
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.ScreenFullPage)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlLogicalText_EmbeddedDocumentCoordinatesRemainScopedThroughPaging(HtmlRenderIntentProfile profile) {
        string html = PositionedLogicalStyles + PositionedLogicalParagraphs("Lead", 4)
            + PositionedLogicalFrame(PositionedLogicalParagraphs("Frame", 6)) + "<p>Tail</p>";
        HtmlRenderOptions options = PositionedLogicalOptions();
        options.PageSize = new OfficePageSize(300D / HtmlRenderOptions.CssPixelsPerInch, 120D / HtmlRenderOptions.CssPixelsPerInch);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(profile, options: options)).Document;

        Assert.Equal("Lead0Lead1Lead2Lead3Frame0Frame1Frame2Frame3Frame4Frame5Tail",
            WithoutPositionedLogicalWhitespace(rendered.Text));
        if (profile == HtmlRenderIntentProfile.ScreenSnapshotPaged) Assert.True(rendered.Pages.Count > 1);
    }

    [Theory]
    [InlineData("section", false)]
    [InlineData("artifact", true)]
    public void HtmlLogicalText_EmbeddedDocumentScopeHonorsExplicitSemantics(string role, bool artifact) {
        string html = PositionedLogicalStyles + PositionedLogicalParagraphs("Lead", 4)
            + PositionedLogicalFrame(PositionedLogicalParagraphs("Frame", 6), "-officeimo-pdf-tag-type:" + role)
            + "<p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        string expected = "Lead0Lead1Lead2Lead3" + (artifact ? string.Empty : "Frame0Frame1Frame2Frame3Frame4Frame5") + "Tail";
        Assert.Equal(expected, WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPositionedText_RetainsIndependentNestedFrameScopes(string position) {
        string inner = PositionedLogicalFrame(PositionedLogicalParagraphs("Inner", 2), height: 60);
        string first = PositionedLogicalFrame("<p>OuterBefore</p>" + inner + "<p>OuterAfter</p>");
        string second = PositionedLogicalFrame("<p>SecondFrame</p>", height: 40);
        string html = PositionedLogicalStyles + PositionedLogicalParagraphs("Lead", 4) + first
            + PositionedLogicalElement(position, "span", "Positioned", 80) + "<p>Between</p>" + second + "<p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("Lead0Lead1Lead2Lead3OuterBeforeInner0Inner1OuterAfterPositionedBetweenSecondFrameTail",
            WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPositionedText_EmbeddedStackingFragmentsRetainOneChildSourceScope(string position) {
        string child = "<p>FrameBefore</p><span style='position:" + position
            + ";left:0;top:80px;z-index:3'>FramePositioned</span><p>FrameAfter</p>";
        string html = PositionedLogicalStyles + PositionedLogicalParagraphs("Lead", 4)
            + PositionedLogicalFrame(child) + "<p>Tail</p>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), PositionedLogicalOptions());

        Assert.Equal("Lead0Lead1Lead2Lead3FrameBeforeFramePositionedFrameAfterTail",
            WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void HtmlLogicalText_IncompleteSourceCoordinatesRetainExistingFallbacks(
        bool firstHasSourceOrder, bool formulaHasLogicalOrder) {
        HtmlRenderText first = PositionedLogicalSyntheticText("First", paintOrder: 1, logicalOrder: 0);
        if (firstHasSourceOrder) first.SemanticSourceOrder = 100;
        HtmlRenderText formulaText = PositionedLogicalSyntheticText("Second", paintOrder: 0,
            logicalOrder: formulaHasLogicalOrder ? 1 : null);
        var formula = new HtmlRenderSemanticGroup(HtmlRenderSemanticGroupRole.Formula,
            0D, 0D, 100D, 20D, new[] { formulaText }, 0, null);
        var rendered = new HtmlRenderDocument(HtmlRenderMode.Continuous,
            new[] { new HtmlRenderPage(1, 300D, 240D, new HtmlRenderVisual[] { formula, first }) },
            new HtmlDiagnosticReport());

        Assert.Equal(formulaHasLogicalOrder ? "FirstSecond" : "SecondFirst",
            WithoutPositionedLogicalWhitespace(rendered.Text));
    }

    [Theory]
    [InlineData(false, "First\nMiddleValue\nLast")]
    [InlineData(true, "First\nMiddle\nValue\nLast")]
    public void HtmlLogicalText_SourceOrderingPreservesReplacementAndArtifactOwnership(
        bool preserveBlockSeparators, string expected) {
        HtmlRenderText first = PositionedLogicalSyntheticText("First", paintOrder: 2, logicalOrder: 0);
        first.SemanticSourceOrder = 10;
        HtmlRenderText middle = PositionedLogicalSyntheticText("Middle", paintOrder: 0, logicalOrder: 1);
        middle.SemanticSourceOrder = 20;
        HtmlRenderText value = PositionedLogicalSyntheticText("Value", paintOrder: 1, logicalOrder: 2);
        value.SemanticSourceOrder = 30;
        HtmlRenderText last = PositionedLogicalSyntheticText("Last", paintOrder: 1, logicalOrder: 3);
        last.SemanticSourceOrder = 40;
        var artifact = new HtmlRenderSemanticGroup(HtmlRenderSemanticGroupRole.Artifact,
            0D, 0D, 100D, 20D, new[] { PositionedLogicalSyntheticText("Ignored", 0, null) }, 2, null);
        var scope = new HtmlRenderLogicalTextScope(preserveBlockSeparators);
        var replacement = new HtmlRenderLogicalTextGroup("Middle\nValue", 0D, 0D, 100D, 40D,
            new HtmlRenderVisual[] { middle, value, artifact }, 0, null, logicalScope: scope);
        var secondary = new HtmlRenderLogicalTextGroup(string.Empty, 0D, 0D, 100D, 40D,
            new[] { middle }, 3, null, logicalScope: scope);
        var rendered = new HtmlRenderDocument(HtmlRenderMode.Continuous,
            new[] { new HtmlRenderPage(1, 300D, 240D, new HtmlRenderVisual[] { replacement, last, first, secondary }) },
            new HtmlDiagnosticReport());

        Assert.Equal(expected, rendered.Text);
    }

    private const string PositionedLogicalStyles = "<style>body,p{margin:0;font:16px/20px Arial}</style>";

    private static string PositionedLogicalElement(string position, string tag, string text, int top) =>
        "<" + tag + " style='position:" + position + ";left:20px;top:" + top
        + "px;width:160px;height:20px'><span>" + text + "</span></" + tag + ">";

    private static string PositionedLogicalParagraphs(string prefix, int count) =>
        string.Concat(Enumerable.Range(0, count).Select(index => "<p>" + prefix + index + "</p>"));

    private static string PositionedLogicalFrame(string content, string style = "", int height = 160) =>
        "<iframe style='display:block;width:240px;height:" + height + "px;" + style + "' srcdoc=\""
        + System.Net.WebUtility.HtmlEncode(PositionedLogicalStyles + content) + "\"></iframe>";

    private static HtmlRenderOptions PositionedLogicalOptions() => new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 300D,
        ViewportHeight = 240D,
        PageSize = new OfficePageSize(300D / HtmlRenderOptions.CssPixelsPerInch, 240D / HtmlRenderOptions.CssPixelsPerInch),
        HonorCssPageRules = false,
        Margins = HtmlRenderMargins.All(0D)
    };

    private static string WithoutPositionedLogicalWhitespace(string text) =>
        string.Concat(text.Where(character => !char.IsWhiteSpace(character)));

    private static HtmlRenderText PositionedLogicalSyntheticText(string text, int paintOrder, int? logicalOrder) =>
        new(text, 0D, 0D, 100D, 20D, new OfficeFontInfo("Arial", 16D), OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, paintOrder, linkUri: null, source: null, semanticRole: null,
            layoutY: null, semanticNodeId: null, textAdvanceWidth: null, logicalTextOrder: logicalOrder);
}
