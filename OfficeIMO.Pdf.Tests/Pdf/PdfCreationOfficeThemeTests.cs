using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCreationOfficeThemeTests {
    [Fact]
    public void DefaultThemeStartsAtCreationAndDoesNotMutateCallerOptions() {
        var options = new PdfCreateOptions();
        Assert.Null(options.DefaultHeadingStyles);
        using var defaultDocument = PdfDocument.Create();
        using var configuredDocument = PdfDocument.Create(options);
        Assert.NotNull(defaultDocument.Options.DefaultHeadingStyles);
        Assert.NotNull(configuredDocument.Options.DefaultHeadingStyles);
        Assert.Null(options.DefaultHeadingStyles);
        defaultDocument.Content.H1("DefaultAnchor");
        using var independent = PdfPigDocument.Open(defaultDocument.ToBytes());
        Assert.Equal(OfficeTheme.Modern.TypeScale.Heading1.ToPoints(), Word(independent, "DefaultAnchor")[0].PointSize, 2);
    }

    [Theory]
    [InlineData("Modern")]
    [InlineData("Classic")]
    [InlineData("Minimal")]
    [InlineData("Boardroom")]
    [InlineData("Dark")]
    public void AuthoredDefaultsReachSavedBodyHeadingsAndTypedTableData(string name) {
        OfficeTheme theme = name switch {
            "Classic" => OfficeTheme.Classic,
            "Minimal" => OfficeTheme.Minimal,
            "Boardroom" => OfficeTheme.Boardroom,
            "Dark" => OfficeTheme.Dark,
            _ => OfficeTheme.Modern
        };
        using var document = PdfDocument.Create(new PdfCreateOptions { Theme = theme });
        document.Content.H1("HeadingAnchor").Text("BodyAnchor")
            .Table(new[] { new { Item = "SavedItem", Amount = 12 } });
        using var output = new MemoryStream();
        document.Save(output);
        byte[] bytes = output.ToArray();
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(bytes, 0, 8));
        using var independent = PdfPigDocument.Open(bytes);
        Letter body = Word(independent, "BodyAnchor")[0];
        Letter heading = Word(independent, "HeadingAnchor")[0];
        Assert.Equal(theme.TypeScale.Body.ToPoints(), body.PointSize, 2);
        Assert.Equal(theme.TypeScale.Heading1.ToPoints(), heading.PointSize, 2);
        AssertColor(body, theme.Palette.Text);
        AssertColor(heading, theme.TextStyle.HeadingColor.Resolve(theme));
        Assert.NotEmpty(Word(independent, "SavedItem"));
        var table = Assert.IsType<TableBlock>(document.Blocks.Last());
        Assert.Equal(new[] { "Item", "Amount" }, table.Rows[0]);
        Assert.Equal(new[] { "SavedItem", "12" }, table.Rows[1]);
        Assert.Equal(theme.TableStyle.HeaderFill.Resolve(theme), table.Style!.HeaderFill);
        Assert.Equal(theme.Palette.Background, document.Options.BackgroundColor);
    }

    [Fact]
    public void ExplicitOptionsAndBlockRunsWinAndCallerMutationDoesNotChangeSavedOutput() {
        var headings = new PdfHeadingStyles { Level1 = new PdfHeadingStyle { FontSize = 17, Color = OfficeColor.Black } };
        var tableStyle = new PdfTableStyle { FontSize = global::OfficeIMO.OfficeLength.Points(13), HeaderFontSize = global::OfficeIMO.OfficeLength.Points(13), HeaderFill = OfficeColor.White };
        var options = new PdfCreateOptions {
            Theme = OfficeTheme.Classic,
            DefaultFont = PdfStandardFont.Courier,
            DefaultFontSize = OfficeLength.Inches(11D / 72D), // Explicitly assigning the legacy value still wins.
            PageWidth = OfficeLength.Inches(6),
            PageHeight = OfficeLength.Inches(8),
            MarginLeft = OfficeLength.Millimeters(12.7),
            MarginRight = OfficeLength.Zero,
            MarginTop = OfficeLength.Centimeters(1.905),
            MarginBottom = OfficeLength.Inches(0.5),
            ShowHeader = true,
            HeaderFormat = "HeaderAnchor",
            HeaderFontSize = OfficeLength.Inches(0.125),
            HeaderOffsetY = OfficeLength.Zero,
            ShowPageNumbers = true,
            FooterFormat = "FooterAnchor",
            FooterFontSize = OfficeLength.Millimeters(3.175),
            FooterOffsetY = OfficeLength.Inches(0.125),
            DefaultTextColor = OfficeColor.FromRgb(12, 34, 56),
            DefaultHeadingStyles = headings,
            DefaultTableStyle = tableStyle,
            DefaultParagraphStyle = new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(3) }
        };
        using var document = PdfDocument.Create(options);
        document.Content.H1("OptionHeading")
            .H1("BlockHeading", new PdfHeadingStyle { FontSize = 19 })
            .Text("OptionBody")
            .Paragraph(paragraph => paragraph.FontSize(7).Color(OfficeColor.Black).Text("RunBody"))
            .Table(new[] { new { Item = "TableBody" } });
        options.Theme = OfficeTheme.Dark;
        options.DefaultFontSize = global::OfficeIMO.OfficeLength.Points(40);
        options.PageWidth = OfficeLength.Inches(9);
        options.Margins = PdfPageMargins.Uniform(OfficeLength.Inches(1));
        options.HeaderOffsetY = OfficeLength.Inches(0.25);
        options.FooterFontSize = OfficeLength.Points(20);
        headings.Level1 = new PdfHeadingStyle { FontSize = 60 };
        tableStyle.FontSize = global::OfficeIMO.OfficeLength.Points(60);
        PdfHeadingStyles detached = document.Options.DefaultHeadingStyles!;
        detached.Level1 = new PdfHeadingStyle { FontSize = 70 };
        using var independent = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(17D, Word(independent, "OptionHeading")[0].PointSize, 2);
        Assert.Equal(19D, Word(independent, "BlockHeading")[0].PointSize, 2);
        Letter body = Word(independent, "OptionBody")[0];
        Assert.Equal(11D, body.PointSize, 2);
        Assert.Contains("Courier", body.FontName, StringComparison.OrdinalIgnoreCase);
        AssertColor(body, OfficeColor.FromRgb(12, 34, 56));
        Assert.Equal(7D, Word(independent, "RunBody")[0].PointSize, 2);
        Assert.Equal(13D, Word(independent, "TableBody")[0].PointSize, 2);
        Assert.Equal(432D, (double)independent.GetPage(1).Width, 2);
        Assert.Equal(576D, (double)independent.GetPage(1).Height, 2);
        Assert.Equal(36D, body.StartBaseLine.X, 2);
        Assert.Equal(OfficeLength.Zero, document.Options.MarginRight);
        Letter header = Word(independent, "HeaderAnchor")[0];
        Letter footer = Word(independent, "FooterAnchor")[0];
        Assert.Equal(9D, header.PointSize, 2);
        Assert.Equal(522D, header.StartBaseLine.Y, 2);
        Assert.Equal(9D, footer.PointSize, 2);
        Assert.Equal(27D, footer.StartBaseLine.Y, 2);
    }

    [Fact]
    public void CloneAndMeasurementViewsRetainInheritedVersusExplicitDefaults() {
        OfficeTheme first = OfficeTheme.Modern.WithTypeScale(OfficeTheme.Modern.TypeScale.With(body: OfficeLength.Points(14)))
            .WithPage(new OfficeThemePage(OfficeLength.Points(21), OfficeLength.Points(22), OfficeLength.Points(23), OfficeLength.Points(24)));
        OfficeTheme second = OfficeTheme.Classic.WithTypeScale(OfficeTheme.Classic.TypeScale.With(body: OfficeLength.Points(16)));
        var inherited = new PdfCreateOptions { Theme = first, MarginRight = global::OfficeIMO.OfficeLength.Points(72) }.CreateDocumentSnapshot(useThemeDefaults: true);
        PdfCreateOptions copy = inherited.Clone();
        _ = copy.DefaultTableStyle; // Populate the derived cache before changing the theme.
        PdfCreateOptions measured = copy.CreateTextMeasurementSnapshot();
        copy.Theme = second;
        Assert.Equal(16D, copy.DefaultFontSize.ToPoints());
        Assert.Equal(second.Page.MarginLeft.ToPoints(), copy.MarginLeft.ToPoints());
        Assert.Equal(72D, copy.MarginRight.ToPoints());
        Assert.Equal(14D, inherited.DefaultFontSize.ToPoints());
        Assert.Equal(14D, measured.DefaultFontSize.ToPoints());
        PdfCreateOptions zeroCopy = copy.Clone();
        zeroCopy.MarginLeft = OfficeLength.Zero;
        PdfCreateOptions zeroMeasured = zeroCopy.CreateTextMeasurementSnapshot();
        zeroCopy = zeroCopy.Clone();
        zeroCopy.Theme = first;
        Assert.Equal(OfficeLength.Zero, zeroCopy.MarginLeft);
        Assert.Equal(OfficeLength.Zero, zeroMeasured.MarginLeft);
        Assert.False(measured.HasSameMeasurementFonts(copy));
        var explicitValues = new PdfCreateOptions {
            Theme = first, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11), FooterFontSize = global::OfficeIMO.OfficeLength.Points(9)
        };
        PdfCreateOptions explicitCopy = explicitValues.CreateDocumentSnapshot(useThemeDefaults: true);
        explicitCopy.Theme = second;
        Assert.Equal(PdfStandardFont.Helvetica, explicitCopy.DefaultFont);
        Assert.Equal(11D, explicitCopy.DefaultFontSize.ToPoints());
        Assert.Equal(9D, explicitCopy.FooterFontSize.ToPoints());
        using var document = PdfDocument.Create(copy);
        document.Content.Text("CloneAnchor");
        using var independent = PdfPigDocument.Open(document.ToBytes());
        Letter anchor = Word(independent, "CloneAnchor")[0];
        Assert.Equal(16D, anchor.PointSize, 2);
        Assert.Equal(second.Page.MarginLeft.ToPoints(), anchor.StartBaseLine.X, 2);
    }

    [Fact]
    public void DeferredTableUsesThemedWeightedColumnsForMeasurementAndPainting() {
        using var document = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Modern });
        document.Content.H1("DeferredAnchor").TableDeferred(() => new[] {
            new[] { "Item", "Quantity" }, new[] { "AlphaRow", "12" }, new[] { "OmegaRow", "3" }
        }, batchSize: 1);
        Assert.True(document.Options.DefaultTableStyle!.AutoFitColumns);
        Assert.False(document.Options.DefaultDeferredTableStyleSnapshot!.AutoFitColumns);
        using var independent = PdfPigDocument.Open(document.ToBytes());
        Assert.NotEmpty(Word(independent, "AlphaRow"));
        Assert.NotEmpty(Word(independent, "OmegaRow"));
        var explicitStyle = new PdfTableStyle { AutoFitColumns = true };
        using var rejected = PdfDocument.Create(new PdfCreateOptions { DefaultTableStyle = explicitStyle });
        rejected.Content.TableDeferred(() => new[] { new[] { "Header" }, new[] { "Value" } });
        Assert.Throws<ArgumentException>(() => rejected.ToBytes());
    }

    [Fact]
    public void FooterThemeAndExplicitFooterOptionsReachSavedText() {
        var theme = OfficeTheme.Modern.WithFooter(new OfficeThemeFooter(
            OfficeLength.Points(8), OfficeColor.FromRgb(10, 20, 30), OfficeThemeRule.None));
        using var inherited = PdfDocument.Create(new PdfCreateOptions { Theme = theme, ShowPageNumbers = true, FooterFormat = "FooterAnchor" });
        inherited.Content.Text("Body");
        using var inheritedPdf = PdfPigDocument.Open(inherited.ToBytes());
        Letter footer = Word(inheritedPdf, "FooterAnchor")[0];
        Assert.Equal(8D, footer.PointSize, 2);
        AssertColor(footer, OfficeColor.FromRgb(10, 20, 30));
        using var explicitDocument = PdfDocument.Create(new PdfCreateOptions {
            Theme = theme, ShowPageNumbers = true, FooterFormat = "ExplicitFooter", FooterFontSize = global::OfficeIMO.OfficeLength.Points(9), FooterTextColor = OfficeColor.Black
        });
        explicitDocument.Content.Text("Body");
        using var explicitPdf = PdfPigDocument.Open(explicitDocument.ToBytes());
        Assert.Equal(9D, Word(explicitPdf, "ExplicitFooter")[0].PointSize, 2);
        AssertColor(Word(explicitPdf, "ExplicitFooter")[0], OfficeColor.Black);
    }

    [Fact]
    public void ImportedLayoutKeepsNeutralDefaultsAndExplicitSourceFormatting() {
        var sourceOptions = new PdfCreateOptions { Theme = OfficeTheme.Dark, MarginRight = global::OfficeIMO.OfficeLength.Points(17), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(15) };
        Assert.Null(sourceOptions.DefaultHeadingStyles);
        Assert.Equal(72D, sourceOptions.MarginLeft.ToPoints());
        using var imported = PdfDocument.CreateForImportedLayout(sourceOptions);
        Assert.Null(imported.Options.DefaultHeadingStyles);
        Assert.Null(imported.Options.DefaultParagraphStyle);
        Assert.Null(imported.Options.DefaultListStyle);
        Assert.Equal(72D, imported.Options.MarginLeft.ToPoints());
        Assert.Equal(17D, imported.Options.MarginRight.ToPoints());
        Assert.Equal(15D, imported.Options.DefaultFontSize.ToPoints());
        Assert.Equal(PdfStandardFont.Helvetica, imported.Options.DefaultFont);
        Assert.Null(imported.Options.DefaultTextColor);
        Assert.Null(imported.Options.EffectiveFooterTextColor);
        imported.Content.Text("ImportedAnchor");
        using var independent = PdfPigDocument.Open(imported.ToBytes());
        Letter body = Word(independent, "ImportedAnchor")[0];
        Assert.Equal(15D, body.PointSize, 2);
        Assert.Equal(72D, body.StartBaseLine.X, 2);
        AssertColor(body, OfficeColor.Black);
        Assert.Same(OfficeTheme.Dark, sourceOptions.Theme);
    }

    [Fact]
    public void TypedTablesUsePageScopedDefaultsAndKeepExplicitNativeStylePrecedence() {
        var scoped = new PdfTableStyle { FontSize = global::OfficeIMO.OfficeLength.Points(14), HeaderFontSize = global::OfficeIMO.OfficeLength.Points(14) };
        var explicitStyle = new PdfTableStyle { FontSize = global::OfficeIMO.OfficeLength.Points(7), HeaderFontSize = global::OfficeIMO.OfficeLength.Points(7) };
        using var document = PdfDocument.Create(builder => builder.Page(page => page
            .DefaultTableStyle(scoped)
            .Content(content => content
                .Table(new[] { new { Item = "PageTableAnchor" } })
                .Table(new[] { new { Item = "ExplicitTableAnchor" } }, new PdfTableOptions { Style = explicitStyle }))));
        scoped.FontSize = global::OfficeIMO.OfficeLength.Points(30);
        explicitStyle.FontSize = global::OfficeIMO.OfficeLength.Points(40);
        using var independent = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(14D, Word(independent, "PageTableAnchor")[0].PointSize, 2);
        Assert.Equal(7D, Word(independent, "ExplicitTableAnchor")[0].PointSize, 2);
    }

    [Fact]
    public async Task LoadInspectAndSaveNeverResetAnIndependentlyProducedPdf() {
        byte[] fixture = IndependentPdf();
        using var input = new MemoryStream(fixture, writable: false);
        using var loaded = await PdfDocument.LoadAsync(input);
        Assert.Equal("SavedIdentity", loaded.Metadata.Title);
        _ = loaded.Inspect();
        Assert.Equal(11D, loaded.Options.DefaultFontSize.ToPoints());
        Assert.Equal(72D, loaded.Options.MarginLeft.ToPoints());
        Assert.Null(loaded.Options.DefaultHeadingStyles);
        loaded.Options.Theme = OfficeTheme.Dark; // Even an internal options read/change must not restyle source bytes.
        Assert.Throws<InvalidOperationException>(() => loaded.Content.Text("Forbidden"));
        using var saved = new MemoryStream();
        loaded.Save(saved);
        Assert.Equal(fixture, saved.ToArray());
        Assert.Equal(fixture, loaded.ToBytes());
        Assert.True(input.CanRead);
        using var independent = PdfPigDocument.Open(saved.ToArray());
        Letter anchor = Word(independent, "SavedAnchor")[0];
        Assert.Equal(25D, anchor.PointSize, 2);
        Assert.Contains("Courier", anchor.FontName, StringComparison.OrdinalIgnoreCase);
        AssertColor(anchor, OfficeColor.FromRgb(255, 0, 0));
    }

    [Fact]
    public void StylingFactoryResolvesRolesAndReturnsIndependentStyleSnapshots() {
        var theme = OfficeTheme.Modern.WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingColor: OfficeColor.Accent1));
        PdfTheme style = PdfTheme.FromOfficeTheme(theme);
        PdfHeadingStyles first = style.HeadingStyles!;
        Assert.Equal(theme.Palette.Accent1, first.Level1!.Color);
        first.Level1 = new PdfHeadingStyle { FontSize = 70 };
        Assert.Equal(theme.TypeScale.Heading1.ToPoints(), style.HeadingStyles!.Level1!.FontSize);
        Assert.Throws<ArgumentNullException>(() => PdfTheme.FromOfficeTheme(null!));
    }

    [Fact]
    public void DarkBackgroundAndExplicitTransparentCanvasStampDefaultsStaySeparate() {
        using var document = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Dark });
        Assert.Equal(OfficeTheme.Dark.Palette.Background, document.Options.BackgroundColor);
        PdfCreateOptions stamp = document.Options.CloneForCanvasStampRendering();
        Assert.Null(stamp.BackgroundColor);
        stamp.Theme = OfficeTheme.Classic;
        Assert.Null(stamp.BackgroundColor);
        using var transparent = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Dark, BackgroundColor = null });
        Assert.Null(transparent.Options.BackgroundColor);
        using var imported = PdfDocument.CreateForImportedLayout(new PdfCreateOptions { Theme = OfficeTheme.Dark });
        Assert.Null(imported.Options.BackgroundColor);
    }

    private static Letter[] Word(PdfPigDocument document, string word) {
        Letter[] letters = document.GetPages().SelectMany(page => page.Letters)
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        string text = string.Concat(letters.Select(letter => letter.Value));
        int start = text.IndexOf(word, StringComparison.Ordinal);
        Assert.True(start >= 0, "Missing saved text: " + word + "; actual: " + text);
        return letters.Skip(start).Take(word.Length).ToArray();
    }

    private static void AssertColor(Letter letter, OfficeColor expected) {
        var (red, green, blue) = letter.Color.ToRGBValues();
        var rgb = expected.ToNormalizedRgb();
        Assert.Equal(rgb.R, red, 3);
        Assert.Equal(rgb.G, green, 3);
        Assert.Equal(rgb.B, blue, 3);
    }

    // Hand-authored independent fixture: no OfficeIMO generation or self-round-trip oracle.
    private static byte[] IndependentPdf() {
        const string content = "BT /F1 25 Tf 1 0 0 rg 20 140 Td (SavedAnchor) Tj ET\n";
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 260 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "endstream",
            "<< /Title (SavedIdentity) >>"
        };
        var text = new StringBuilder("%PDF-1.4\n");
        var offsets = new int[objects.Length + 1];
        for (int index = 0; index < objects.Length; index++) {
            offsets[index + 1] = Encoding.ASCII.GetByteCount(text.ToString());
            text.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = Encoding.ASCII.GetByteCount(text.ToString());
        text.Append("xref\n0 7\n0000000000 65535 f \n");
        foreach (int offset in offsets.Skip(1)) text.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        text.Append("trailer\n<< /Size 7 /Root 1 0 R /Info 6 0 R >>\nstartxref\n")
            .Append(xref.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(text.ToString());
    }
}
