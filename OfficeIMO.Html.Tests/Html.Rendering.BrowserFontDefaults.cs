using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(true, "<pre>X</pre>", "Consolas")]
    [InlineData(true, "<code>X</code>", "Consolas")]
    [InlineData(true, "<kbd>X</kbd>", "Consolas")]
    [InlineData(true, "<samp>X</samp>", "Consolas")]
    [InlineData(true, "<pre><code><samp>X</samp></code></pre>", "Consolas")]
    // The user-agent code rule beats an inherited family in both modes, as in a browser.
    [InlineData(true, "<section style='font-family:serif'><code>X</code></section>", "Consolas")]
    [InlineData(false, "<code>X</code>", "Consolas")]
    public void HtmlBrowserFontDefaults_CodeFamilyIsAppliedBeforeChildInheritance(
        bool standardsMode, string markup, string documentFamily) {
        AssertBrowserFontDefaultModes(markup, 13D, "monospace", 16D, documentFamily,
            standardsMode: standardsMode);
    }

    [Theory]
    [InlineData("<code style='font-size:1em'>X</code>", 13D, 16D)]
    [InlineData("<code style='font-size:inherit'>X</code>", 13D, 16D)]
    [InlineData("<section style='font-size:20px'><code>X</code></section>", 20D, 20D)]
    [InlineData("<section style='font-size:1.25em'><code>X</code></section>", 16.25D, 20D)]
    [InlineData("<section style='font-size:20px'><pre style='font-size:inherit'>X</pre></section>", 20D, 20D)]
    [InlineData("<section style='font-size:20px'><code style='font-size:1.25em'>X</code></section>", 25D, 25D)]
    [InlineData("<section style='font-size:20px'><kbd style='font-size:125%'>X</kbd></section>", 25D, 25D)]
    [InlineData("<section style='font-size:20px'><pre style='font-size:medium'>X</pre></section>", 13D, 16D)]
    [InlineData("<section style='font-size:20px'><code style='font-size:initial'>X</code></section>", 13D, 16D)]
    [InlineData("<section style='font-size:20px'><kbd style='font-size:revert'>X</kbd></section>", 20D, 20D)]
    [InlineData("<pre style='font-size:18px'><code style='font-size:1em'><samp>X</samp></code></pre>", 18D, 18D)]
    public void HtmlBrowserFontDefaults_UsedSizeRetainsAbsoluteOrRelativeOrigin(
        string markup, double browserSize, double documentSize) {
        AssertBrowserFontDefaultModes(markup, browserSize, "monospace", documentSize, "Consolas");
    }

    [Theory]
    [InlineData("<code style='font-family:inherit'>X</code>", 20D, "serif", 20D, "serif")]
    [InlineData("<code style='font-family:unset'>X</code>", 20D, "serif", 20D, "serif")]
    [InlineData("<samp style='font-family:initial'>X</samp>", 20D, "serif", 20D, "Consolas")]
    [InlineData("<samp style='font:18px monospace'>X</samp>", 18D, "monospace", 18D, "monospace")]
    public void HtmlBrowserFontDefaults_AuthoredFamilyAndFontShorthandOverrideTheDefault(
        string markup, double browserSize, string browserFamily, double documentSize, string documentFamily) {
        AssertBrowserFontDefaultModes("<section style='font-size:20px;font-family:serif'>" + markup + "</section>",
            browserSize, browserFamily, documentSize, documentFamily);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("body")]
    public void HtmlBrowserFontDefaults_RootAndBodyGenericTransitionsRestoreTheCurrentMedium(string owner) {
        const string markup = "<div>MONO</div><section style='font-family:serif'>"
            + "<div>SERIF</div><code>CODE</code></section>";
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser,
                extraCss: owner + "{font-family:monospace}");
            HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>().ToArray();
            HtmlRenderText mono = Assert.Single(text, item => item.Text == "MONO");
            HtmlRenderText serif = Assert.Single(text, item => item.Text == "SERIF");
            HtmlRenderText code = Assert.Single(text, item => item.Text == "CODE");
            Assert.Equal(browser ? 13D : 16D, mono.Font.Size, 6);
            Assert.Equal("monospace", mono.Font.FamilyName);
            Assert.Equal(16D, serif.Font.Size, 6);
            Assert.Equal("serif", serif.Font.FamilyName);
            Assert.Equal(browser ? 13D : 16D, code.Font.Size, 6);
            Assert.Equal(browser ? "monospace" : "Consolas", code.Font.FamilyName);
        }
    }

    [Theory]
    [InlineData(14D)]
    [InlineData(16D)]
    [InlineData(20D)]
    public void HtmlBrowserFontDefaults_CustomStandardMediumKeepsTheFixedMediumIndependent(double standardSize) {
        const string markup = "<div>ORDINARY</div><div><code>CODE</code></div>"
            + "<div><code style='font-size:medium'>MEDIUM</code></div>"
            + "<section style='font-size:1.25em'><code>RELATIVE</code></section>"
            + "<section style='font-size:20px'><code>ABSOLUTE</code></section>";
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser,
                defaultFontSize: standardSize);
            HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderText>().ToArray();
            Assert.Equal(standardSize, Assert.Single(text, item => item.Text == "ORDINARY").Font.Size, 6);
            foreach (string marker in new[] { "CODE", "MEDIUM" }) {
                HtmlRenderText run = Assert.Single(text, item => item.Text == marker);
                Assert.Equal(browser ? 13D : standardSize, run.Font.Size, 6);
                Assert.Equal(browser ? "monospace" : "Consolas", run.Font.FamilyName);
            }
            HtmlRenderText relative = Assert.Single(text, item => item.Text == "RELATIVE");
            Assert.Equal(browser ? 16.25D : standardSize * 1.25D, relative.Font.Size, 6);
            Assert.Equal(browser ? "monospace" : "Consolas", relative.Font.FamilyName);
            Assert.Equal(20D, Assert.Single(text, item => item.Text == "ABSOLUTE").Font.Size, 6);
        }
    }

    [Theory]
    [InlineData("monospace", 13D)]
    [InlineData("/* fixed */ m\\6f nospace /* tail */", 13D)]
    [InlineData("'monospace'", 16D)]
    [InlineData("monospace,serif", 16D)]
    [InlineData("'Courier New',monospace", 16D)]
    public void HtmlBrowserFontDefaults_OnlyASoleDecodedGenericSelectsTheFixedMedium(
        string family, double browserSize) {
        foreach (bool browser in new[] { false, true }) {
            string markup = "<section style=\"font-family:" + family
                + "\"><div>PARENT</div><span>CHILD</span></section>";
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser);
            HtmlRenderText[] text = BrowserFontDefaultText(rendered);
            HtmlRenderText parent = Assert.Single(text, item => item.Text == "PARENT");
            HtmlRenderText child = Assert.Single(text, item => item.Text == "CHILD");
            Assert.Equal(browser ? browserSize : 16D, parent.Font.Size, 6);
            Assert.Equal(parent.Font.Size, child.Font.Size, 6);
            Assert.Equal(parent.Font.FamilyName, child.Font.FamilyName);
        }
    }

    [Theory]
    [InlineData(true, 16D, "small", 13D, 12D, 14.24D)]
    [InlineData(true, 16D, "large", 18D, 16D, 19.2D)]
    [InlineData(true, 16D, "xxx-large", 48D, 39D, 48D)]
    [InlineData(false, 16D, "xxx-large", 48D, 40D, 48D)]
    [InlineData(true, 14D, "x-small", 10D, 10D, 10.5D)]
    [InlineData(false, 14D, "x-small", 9D, 9D, 10.5D)]
    [InlineData(true, 9D, "xx-small", 9D, 9D, 5.4D)]
    [InlineData(true, 20D, "small", 17.8D, 12D, 17.8D)]
    public void HtmlBrowserFontDefaults_AbsoluteKeywordsAreReevaluatedForTheCurrentGeneric(
        bool standardsMode, double standardSize, string keyword,
        double browserParentSize, double browserChildSize, double documentSize) {
        string markup = "<section style='font-size:" + keyword
            + "'><div>PARENT</div><code>CHILD</code></section>";
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser,
                standardsMode, defaultFontSize: standardSize);
            HtmlRenderText[] text = BrowserFontDefaultText(rendered);
            HtmlRenderText parent = Assert.Single(text, item => item.Text == "PARENT");
            HtmlRenderText child = Assert.Single(text, item => item.Text == "CHILD");
            Assert.Equal(browser ? browserParentSize : documentSize, parent.Font.Size, 6);
            Assert.Equal(browser ? browserChildSize : documentSize, child.Font.Size, 6);
            Assert.Equal(browser ? "serif" : "Arial", parent.Font.FamilyName);
            Assert.Equal(browser ? "monospace" : "Consolas", child.Font.FamilyName);
        }
    }

    [Theory]
    [InlineData("<h1><code>X</code></h1>", 26D, 32D)]
    [InlineData("<small><code>X</code></small>", 10.833333333333334D, 13.28D)]
    [InlineData("<big><code>X</code></big>", 15.6D, 18.72D)]
    [InlineData("<section style='font-size:20px'><big><code>X</code></big></section>", 24D, 23.4D)]
    [InlineData("<span style='font-size:larger'><code>X</code></span>", 15.6D, 19.2D)]
    [InlineData("<span style='font-size:smaller'><code>X</code></span>", 10.833333333333334D, 13.333333333333334D)]
    public void HtmlBrowserFontDefaults_UaAndRelativeTransformsRetainTheirUsedScale(
        string markup, double browserSize, double documentSize) {
        AssertBrowserFontDefaultModes(markup, browserSize, "monospace", documentSize, "Consolas");
    }

    [Theory]
    [InlineData("1.25rem", 16.25D)]
    [InlineData("calc(10px + 10px)", 20D)]
    [InlineData("calc(1em + .25em)", 20D)]
    [InlineData("calc(10px + .625em)", 20D)]
    public void HtmlBrowserFontDefaults_RemAndCalculatedSizesRetainAbsoluteProvenance(
        string size, double browserSize) {
        AssertBrowserFontDefaultModes("<section style='font-size:" + size
            + "'><code>X</code></section>", browserSize, "monospace", 20D, "Consolas",
            extraCss: "html{font-family:monospace}body{font-family:serif}");
    }

    [Fact]
    public void HtmlBrowserFontDefaults_RootRemAndContainerQueriesUseTheSameUsedFontSize() {
        const string css = "html{font-family:monospace}body{font-family:serif}"
            + "#font-container{font-family:monospace;width:10em;container-type:inline-size;background:#eee}"
            + "#container-width-target,#container-em-target{width:1px;height:1em;background:red}"
            + "@container (width:130px){#container-width-target{width:20px}}"
            + "@container (width>9em){#container-em-target{width:20px}}";
        const string markup = "<div id='root-rem-box' style='width:10rem;height:1rem;background:blue'></div>"
            + "<section id='font-container'><div id='container-width-target'></div>"
            + "<div id='container-em-target'></div></section>";
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser, extraCss: css);
            HtmlRenderShape rem = FindShape(rendered, "div#root-rem-box");
            HtmlRenderShape container = FindShape(rendered, "section#font-container");
            HtmlRenderShape absoluteQuery = FindShape(rendered, "div#container-width-target");
            HtmlRenderShape emQuery = FindShape(rendered, "div#container-em-target");
            double medium = browser ? 13D : 16D;
            Assert.Equal(10D * medium, rem.Width, 6);
            Assert.Equal(medium, rem.Height, 6);
            Assert.Equal(10D * medium, container.Width, 6);
            Assert.Equal(2D * medium, container.Height, 6);
            Assert.Equal(browser ? 20D : 1D, absoluteQuery.Width, 6);
            Assert.Equal(medium, absoluteQuery.Height, 6);
            Assert.Equal(20D, emQuery.Width, 6);
            Assert.Equal(medium, emQuery.Height, 6);
        }
    }

    private static HtmlRenderText[] BrowserFontDefaultText(HtmlRenderDocument rendered) =>
        rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>().ToArray();

    [Theory]
    [InlineData("table")]
    [InlineData("table-row")]
    [InlineData("table-cell")]
    public void HtmlBrowserFontProvenance_AnonymousTableWrappersRetainRelativeInheritance(string display) {
        AssertBrowserFontDefaultModes("<div style='display:" + display + "'><code>X</code></div>",
            13D, "monospace", 16D, "Consolas");
    }

    [Theory]
    [InlineData("", 8D, 6.5D)]
    [InlineData("font-size:20px", 10D, 10D)]
    public void HtmlBrowserFontProvenance_RubyAnnotationDescendantsInheritTheReducedSize(
        string style, double expectedSize, double browserCodeSize) {
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture("<section style='" + style
                + "'><ruby>A<rt><span>X</span><code>Z</code></rt></ruby><ruby>B<rt>Y</rt></ruby></section>", browser);
            HtmlRenderText[] text = BrowserFontDefaultText(rendered);
            Assert.Equal(expectedSize, Assert.Single(text, run => run.Text == "X").Font.Size, 6);
            Assert.Equal(Assert.Single(text, run => run.Text == "Y").Font.Size,
                Assert.Single(text, run => run.Text == "X").Font.Size, 6);
            Assert.Equal(browser ? browserCodeSize : expectedSize,
                Assert.Single(text, run => run.Text == "Z").Font.Size, 6);
        }
    }

    [Theory]
    [InlineData("font-family:monospace", "", 6.5D, 8D)]
    [InlineData("font-family:monospace;font-size:inherit", "", 13D, 16D)]
    [InlineData("font-size:inherit", "", 16D, 16D)]
    [InlineData("font-size:unset", "", 16D, 16D)]
    [InlineData("font-size:initial", "", 16D, 16D)]
    [InlineData("font-size:revert", "", 8D, 8D)]
    [InlineData("font-family:monospace", "font-size:20px", 10D, 10D)]
    [InlineData("font-size:12px", "", 12D, 12D)]
    public void HtmlBrowserFontProvenance_RubyDefaultsRespectAuthoredFamilyAndSize(
        string annotationCss, string outerCss, double browserSize, double documentSize) {
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture("<section style='" + outerCss
                + "'><ruby>A<rt style='" + annotationCss + "'>D<span>N</span></rt></ruby></section>", browser);
            HtmlRenderText[] text = BrowserFontDefaultText(rendered);
            HtmlRenderText[] annotationRuns = text.Where(run => run.Text != "A").ToArray();
            Assert.Equal("DN", string.Concat(annotationRuns.Select(run => run.Text)));
            Assert.All(annotationRuns, run => Assert.Equal(browser ? browserSize : documentSize, run.Font.Size, 6));
        }
    }

    [Theory]
    [InlineData("font-family:inherit")]
    [InlineData("font-family:unset")]
    [InlineData("font:inherit")]
    [InlineData("font-family:var(--missing,inherit)")]
    public void HtmlBrowserFontProvenance_ExplicitFamilyInheritanceUsesTheUnstyledParent(string declaration) {
        AssertBrowserFontDefaultModes("<code style='" + declaration + "'>X</code>",
            16D, "serif", 16D, "Arial");
    }

    [Theory]
    [InlineData("initial", 16D, "serif")]
    [InlineData("revert", 13D, "monospace")]
    public void HtmlBrowserFontProvenance_FamilyResetAndRollbackRetainTheirPolicies(
        string keyword, double browserSize, string browserFamily) {
        AssertBrowserFontDefaultModes("<code style='font-family:" + keyword + "'>X</code>",
            browserSize, browserFamily, 16D, "Consolas");
    }

    [Fact]
    public void HtmlBrowserFontProvenance_MathContainerUsesItsEffectiveMathFamily() {
        const string markup = "<section style='font-family:monospace'><math id='math-default' "
            + "style='width:10em;container-type:inline-size;background:blue'>"
            + "<mtext id='math-token'>x</mtext></math></section>";
        const string css = "@container (width:160px){#math-token{text-transform:uppercase}}";
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser, extraCss: css);
            HtmlRenderShape[] shapes = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
                .OfType<HtmlRenderShape>().Where(shape => shape.Source == "math#math-default").ToArray();
            Assert.NotEmpty(shapes);
            Assert.All(shapes, shape => Assert.Equal(160D, shape.Width, 6));
            HtmlRenderDrawing drawing = Assert.Single(rendered.Pages
                .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderDrawing>());
            OfficeIMO.Drawing.OfficeDrawingText text = Assert.Single(drawing.Drawing.Elements
                .OfType<OfficeIMO.Drawing.OfficeDrawingText>());
            Assert.Equal("X", text.RasterText ?? text.Text);
        }
    }

    private static void AssertBrowserFontDefaultModes(
        string markup, double browserSize, string browserFamily, double documentSize, string documentFamily,
        bool standardsMode = true, string extraCss = "") {
        foreach (bool browser in new[] { false, true }) {
            HtmlRenderDocument rendered = RenderBrowserFontDefaultFixture(markup, browser, standardsMode, extraCss);
            HtmlRenderText text = Assert.Single(rendered.Pages
                .SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderText>());
            Assert.Equal("X", text.Text);
            Assert.Equal(browser ? browserSize : documentSize, text.Font.Size, 6);
            Assert.Equal(browser ? browserFamily : documentFamily, text.Font.FamilyName);
        }
    }

    private static HtmlRenderDocument RenderBrowserFontDefaultFixture(
        string markup, bool browser, bool standardsMode = true, string extraCss = "", double defaultFontSize = 16D) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 360D,
            ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D),
            AllowSystemFontFallback = false,
            DefaultFontSize = defaultFontSize
        };
        if (browser) options.UseBrowserUserAgentStyles();
        string html = (standardsMode ? "<!doctype html>" : "")
            + "<html><head><style>body{margin:0}" + extraCss + "</style></head><body>"
            + markup + "</body></html>";
        return HtmlRenderTestDriver.Render(html, options);
    }
}
