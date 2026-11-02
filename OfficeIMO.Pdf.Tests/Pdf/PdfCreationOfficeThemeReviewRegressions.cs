using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCreationOfficeThemeReviewRegressions {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EagerDeferredAndContextTablesKeepTheirOwningPageDefaults(bool pageTheme) {
        using var document = PdfDocument.Create(new PdfCreateOptions { DefaultTableStyle = TableStyle(9) });
        document.Compose(builder => builder
            .Page(page => {
                SetPageStyle(page, 14, pageTheme);
                page.Content(content => TablePaths(content, "First"));
            })
            .Page(page => {
                SetPageStyle(page, 16, pageTheme);
                page.Content(content => TablePaths(content, "Second"));
            }));
        AddTable(document.Content, "RootAfterPages");

        using var saved = PdfPigDocument.Open(Save(document));
        AssertPathSizes(saved.GetPage(1).Letters, "First", 14);
        AssertPathSizes(saved.GetPage(2).Letters, "Second", 16);
        AssertSize(saved.GetPages().SelectMany(page => page.Letters), "RootAfterPages", 9);
    }

    [Fact]
    public void DeferredDefaultsFollowTheOwningPageAfterRegistrationAndRestoreAfterFailure() {
        bool reject = true;
        using var document = PdfDocument.Create(new PdfCreateOptions { DefaultTableStyle = TableStyle(9) });
        document.Compose(builder => builder.Page(page => page
            .Content(content => content.Deferred(_ => {
                if (reject) throw new InvalidOperationException("Callback failed");
                return deferred => AddTable(deferred, "DeferredAfterFailure");
            }))
            .DefaultTableStyle(TableStyle(14))));
        Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        AddTable(document.Content, "RootAfterFailure");
        reject = false;

        using var saved = PdfPigDocument.Open(Save(document));
        IEnumerable<Letter> letters = saved.GetPages().SelectMany(page => page.Letters);
        AssertSize(letters, "DeferredAfterFailure", 14);
        AssertSize(letters, "RootAfterFailure", 9);
    }

    [Fact]
    public void ExplicitStandardFontsOverrideInheritedBoardroomHeadingsAtDocumentAndPageScope() {
        using var document = PdfDocument.Create(new PdfCreateOptions {
            Theme = OfficeTheme.Boardroom, DefaultFont = PdfStandardFont.Courier
        });
        AddHeadings(document.Content, "DocumentBase");
        document.Compose(builder => builder
            .Page(page => page.DefaultTextStyle(style => style.Font(PdfStandardFont.Helvetica))
                .Content(content => {
                    AddHeadings(content, "First");
                    content.Deferred(_ => deferred => AddHeadings(deferred, "FirstDeferred"));
                }))
            .Page(page => page.DefaultTextStyle(style => style.Font(PdfStandardFont.Courier))
                .Content(content => AddHeadings(content, "Second"))));
        AddHeadings(document.Content, "DocumentTail");

        using var saved = PdfPigDocument.Open(Save(document));
        AssertHeadingFonts(saved, "DocumentBase", "Courier");
        AssertHeadingFonts(saved, "First", "Helvetica");
        AssertHeadingFonts(saved, "FirstDeferred", "Helvetica");
        AssertHeadingFonts(saved, "Second", "Courier");
        AssertHeadingFonts(saved, "DocumentTail", "Courier");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CallerEmbeddedFamilyOverridesInheritedBoardroomHeadingFonts(bool pageScope, bool cff) {
        string path = Assert.IsType<string>(cff
            ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont()
            : PdfComplianceTestFonts.FindBundledTrueTypeFont());
        byte[] font = File.ReadAllBytes(path);
        var family = new PdfEmbeddedFontFamily("ThemeCaller", font, bold: font, italic: font, boldItalic: font);
        var options = new PdfCreateOptions { Theme = OfficeTheme.Boardroom };
        if (!pageScope) options.UseFontFamily(family);
        using var document = PdfDocument.Create(options);
        if (pageScope) {
            document.Compose(builder => builder
                .Page(page => page.UseFontFamily(family).Content(content => {
                    AddHeadings(content, "Embedded");
                    content.Component(new ContextHeadings("EmbeddedContext"));
                }))
                .Page(page => page.DefaultTextStyle(style => style.Font(PdfStandardFont.Courier))
                    .Content(content => AddHeadings(content, "Second"))));
        } else {
            AddHeadings(document.Content, "Embedded");
            document.Content.Component(new ContextHeadings("EmbeddedContext"));
        }

        using var saved = PdfPigDocument.Open(Save(document));
        AssertHeadingFonts(saved, "Embedded", "ThemeCaller");
        AssertHeadingFonts(saved, "EmbeddedContext", "ThemeCaller");
        if (pageScope) AssertHeadingFonts(saved, "Second", "Courier");
    }

    [Fact]
    public void ExplicitHeadingStylesAndRunsStayStrongerAndTheStandaloneFactoryKeepsDistinctFonts() {
        PdfTheme boardroom = PdfTheme.FromOfficeTheme(OfficeTheme.Boardroom);
        using var document = PdfDocument.Create(new PdfCreateOptions {
            Theme = OfficeTheme.Boardroom, DefaultFont = PdfStandardFont.Courier,
            DefaultHeadingStyles = boardroom.HeadingStyles
        });
        document.Content.H1("ExplicitDefaultHeading")
            .H2("ExplicitBlockHeading", new PdfHeadingStyle { Font = PdfStandardFont.Helvetica })
            .Heading(3, "ExplicitRunHeading", paragraph => paragraph.Font(PdfStandardFont.Courier).Text("ExplicitRunHeading"),
                PdfAlign.Left, null, new PdfHeadingStyle { Font = PdfStandardFont.TimesRoman })
            .Text("ExplicitBody");
        document.Compose(builder => builder.Page(page => page.Theme(boardroom)
            .DefaultTextStyle(style => style.Font(PdfStandardFont.Courier))
            .Content(content => content.H1("ExplicitPageThemeHeading").Text("ExplicitPageBody"))));

        using var saved = PdfPigDocument.Open(Save(document));
        AssertFont(saved, "ExplicitDefaultHeading", "Times");
        AssertFont(saved, "ExplicitBlockHeading", "Helvetica");
        AssertFont(saved, "ExplicitRunHeading", "Courier");
        AssertFont(saved, "ExplicitBody", "Courier");
        AssertFont(saved, "ExplicitPageThemeHeading", "Times");
        AssertFont(saved, "ExplicitPageBody", "Courier");

        using var inherited = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Boardroom });
        inherited.Content.H1("InheritedBoardroomHeading");
        using var inheritedSaved = PdfPigDocument.Open(Save(inherited));
        AssertFont(inheritedSaved, "InheritedBoardroomHeading", "Times");
    }

    [Fact]
    public void LateExplicitFontAndClonesKeepOnlyAuthoredPerLevelHeadingFonts() {
        using var document = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Boardroom });
        _ = document.Options.DefaultHeadingStyles;
        document.Options.SetDefaultHeadingStyle(2, new PdfHeadingStyle { Font = PdfStandardFont.TimesRoman });
        document.Options.DefaultFont = PdfStandardFont.Courier;
        AddHeadings(document.Content, "Late");
        using var cloned = PdfDocument.Create(document.Options.Clone());
        AddHeadings(cloned.Content, "Clone");

        using var saved = PdfPigDocument.Open(Save(document));
        using var cloneSaved = PdfPigDocument.Open(Save(cloned));
        AssertFont(saved, "LateOne", "Courier");
        AssertFont(saved, "LateTwo", "Times");
        AssertFont(saved, "LateThree", "Courier");
        AssertFont(cloneSaved, "CloneOne", "Courier");
        AssertFont(cloneSaved, "CloneTwo", "Times");
        AssertFont(cloneSaved, "CloneThree", "Courier");
    }

    private static PdfTableStyle TableStyle(double size) => new PdfTableStyle {
        FontSize = global::OfficeIMO.OfficeLength.Points(size), HeaderFontSize = global::OfficeIMO.OfficeLength.Points(size), SpacingAfter = global::OfficeIMO.OfficeLength.Points(3)
    };

    private static void SetPageStyle(PdfPageBuilder page, double size, bool pageTheme) {
        if (pageTheme) page.Theme(new PdfTheme { TableStyle = TableStyle(size) });
        else page.DefaultTableStyle(TableStyle(size));
    }

    private static void TablePaths(PdfContentBuilder content, string prefix) {
        AddTable(content, prefix + "Eager");
        content.Component(new EagerTable(prefix + "Component"))
            .Deferred(_ => deferred => AddTable(deferred, prefix + "Deferred"))
            .Component(new ContextTable(prefix + "Context"))
            .Deferred(_ => outer => outer.Deferred(nestedContext => inner => AddTable(inner, prefix + "Nested")));
    }

    private static void AddTable(PdfContentBuilder content, string text) => content.Table(new[] { new { Item = text } });

    private static void AddHeadings(PdfContentBuilder content, string prefix) => content
        .H1(prefix + "One").H2(prefix + "Two").H3(prefix + "Three").Text(prefix + "Body");

    private static byte[] Save(PdfDocument document) {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    private static void AssertPathSizes(IEnumerable<Letter> letters, string prefix, double size) {
        foreach (string suffix in new[] { "Eager", "Component", "Deferred", "Context", "Nested" })
            AssertSize(letters, prefix + suffix, size);
    }

    private static void AssertSize(IEnumerable<Letter> letters, string text, double size) =>
        Assert.All(Word(letters, text), letter => Assert.Equal(size, letter.PointSize, 2));

    private static void AssertHeadingFonts(PdfPigDocument document, string prefix, string family) {
        foreach (string suffix in new[] { "One", "Two", "Three", "Body" }) AssertFont(document, prefix + suffix, family);
    }

    private static void AssertFont(PdfPigDocument document, string text, string family) =>
        Assert.All(Word(document.GetPages().SelectMany(page => page.Letters), text),
            letter => Assert.Contains(family, letter.FontName, StringComparison.OrdinalIgnoreCase));

    private static Letter[] Word(IEnumerable<Letter> source, string word) {
        Letter[] letters = source.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        Assert.True(index >= 0, "Missing saved text: " + word + "; actual: " + text);
        return letters.Skip(index).Take(word.Length).ToArray();
    }

    private sealed class EagerTable : IPdfComponent {
        private readonly string _text;
        internal EagerTable(string text) { _text = text; }
        public void Compose(PdfContentBuilder content) => AddTable(content, _text);
    }

    private sealed class ContextTable : IPdfContextComponent {
        private readonly string _text;
        internal ContextTable(string text) { _text = text; }
        public void Compose(PdfContentBuilder content, PdfFlowContext context) => AddTable(content, _text);
    }

    private sealed class ContextHeadings : IPdfContextComponent {
        private readonly string _prefix;
        internal ContextHeadings(string prefix) { _prefix = prefix; }
        public void Compose(PdfContentBuilder content, PdfFlowContext context) => AddHeadings(content, _prefix);
    }
}
