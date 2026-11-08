using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void Markdown_OfficeThemeProjectsNativeTypographyAndKeepsRichHeadingRoles() {
        OfficeTheme native = OfficeTheme.Modern
            .WithFonts(OfficeTheme.Modern.Fonts.With(
                body: OfficeTheme.Modern.Fonts.Body.With(family: "Times New Roman"),
                heading: OfficeTheme.Modern.Fonts.Heading.With(family: "Courier New")))
            .WithTypeScale(OfficeTheme.Modern.TypeScale.With(heading1: OfficeLength.Points(29),
                heading4: OfficeLength.Points(19), body: OfficeLength.Points(14)))
            .WithTextStyle(OfficeTheme.Modern.TextStyle.With(headingColor: OfficeColor.Parse("#800000"),
                linkColor: OfficeColor.Parse("#008000"), headingBold: false))
            .WithTableStyle(OfficeTheme.Modern.TableStyle.With(headerBold: false, headerFill: OfficeColor.Parse("#CCE8F5")))
            .WithSpacing(OfficeTheme.Modern.Spacing.With(paragraphAfter: OfficeLength.Points(21)))
            .WithPage(OfficeTheme.Modern.Page.With(marginLeft: OfficeLength.Points(37)));
        using var source = MarkdownDocument.Parse("# Heading *italic* [link](https://example.test)\n\nBodyone\n\nBodytwo\n\n#### Deep **bold**\n\n| Header | Count |\n| --- | --- |\n| Tablecell | 2 |\n");
        var options = new MarkdownPdfExportOptions {
            Theme = MarkdownVisualTheme.FromOfficeTheme(native),
            PdfOptions = new PdfCore.PdfCreateOptions().EnableTaggedPdfCatalogMarkers(),
            ShowPageNumbers = false,
            TextFallbacks = PdfCore.PdfTextFallbackFeatures.None,
            ResourcePolicy = new PdfCore.PdfResourcePolicy { AllowSystemFontEmbedding = false }
        };
        using PdfCore.PdfDocument model = source.ToPdf(options);
        byte[] bytes = model.ToBytes();
        using var artifact = PdfPigDocument.Open(bytes);
        var words = artifact.GetPage(1).GetWords().ToList();
        var heading = Assert.Single(words, word => word.Text == "Heading");
        var body = Assert.Single(words, word => word.Text == "Bodyone");
        var next = Assert.Single(words, word => word.Text == "Bodytwo");
        var deep = Assert.Single(words, word => word.Text == "Deep");
        Assert.Equal(29D, heading.Letters[0].PointSize, 2);
        Assert.Contains("Courier", heading.Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Bold", heading.Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(14D, body.Letters[0].PointSize, 2);
        Assert.Contains("Times", body.Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.InRange(body.Letters[0].StartBaseLine.X, 36.9, 37.1);
        Assert.True(body.Letters[0].StartBaseLine.Y - next.Letters[0].StartBaseLine.Y > 32);
        Assert.Equal(19D, deep.Letters[0].PointSize, 2);
        Assert.Contains("Bold", words.Single(word => word.Text == "bold").Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Oblique", words.Single(word => word.Text == "italic").Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        var outline = Assert.Single(PdfCore.PdfInspector.Inspect(bytes).Outlines);
        Assert.Equal("Deep bold", Assert.Single(outline.Children).Title);
        using var logical = PdfCore.PdfDocument.Load(bytes);
        Assert.Single(logical.Read().GetLinksByUri("https://example.test/"));
        Assert.Equal(4, Assert.Single(logical.Read().Headings, heading => heading.Text == "Deep bold").Level);
        string content = PdfTestContentText.ReadDecoded(bytes);
        Assert.Contains("0.502 0 0 rg", content, StringComparison.Ordinal);
        Assert.Contains("0 0.502 0 rg", content, StringComparison.Ordinal);

        MarkdownVisualTheme customized = MarkdownVisualTheme.FromOfficeTheme(native).WithColors(heading: "#000080");
        MarkdownVisualPalette editedPalette = customized.Palette.Clone();
        editedPalette.TableHeaderBackground = OfficeColor.Blue;
        customized.Palette = editedPalette;
        customized.Table = new MarkdownTableVisualStyle { CellPaddingX = OfficeLength.Points(13), CellPaddingY = OfficeLength.Points(8),
            BorderWidth = native.TableStyle.BorderWidth, EmphasizeHeader = false, UseRowStripes = false };
        options.Theme = customized;
        using PdfCore.PdfDocument edited = source.ToPdf(options);
        byte[] editedBytes = edited.ToBytes();
        using var editedArtifact = PdfPigDocument.Open(editedBytes);
        var editedWords = editedArtifact.GetPage(1).GetWords().ToList();
        Assert.Contains("Courier", Assert.Single(editedWords, word => word.Text == "Heading").Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Bold", Assert.Single(editedWords, word => word.Text == "Header").Letters[0].FontName, StringComparison.OrdinalIgnoreCase);
        double originalCellX = Assert.Single(words, word => word.Text == "Tablecell").Letters[0].StartBaseLine.X;
        double editedCellX = Assert.Single(editedWords, word => word.Text == "Tablecell").Letters[0].StartBaseLine.X;
        Assert.InRange(editedCellX - originalCellX, 5.7, 6.3);
        Assert.InRange(Assert.Single(editedWords, word => word.Text == "Bodyone").Letters[0].StartBaseLine.X, 36.9, 37.1);
        string editedContent = PdfTestContentText.ReadDecoded(editedBytes);
        Assert.Contains("0 0 0.502 rg", editedContent, StringComparison.Ordinal);
        Assert.Contains("0 0 1 rg", editedContent, StringComparison.Ordinal);
        Assert.Contains("0 0.502 0 rg", editedContent, StringComparison.Ordinal);

        options.PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultFont = PdfCore.PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(9),
            DefaultHeadingStyles = new PdfCore.PdfHeadingStyles { Level1 = new PdfCore.PdfHeadingStyle { FontSize = OfficeLength.Points(17) } },
            PageSetup = { Margins = { Left = OfficeLength.Points(83) } }
        };
        using PdfCore.PdfDocument overridden = source.ToPdf(options);
        using var explicitArtifact = PdfPigDocument.Open(overridden.ToBytes());
        var explicitWords = explicitArtifact.GetPage(1).GetWords().ToList();
        Assert.Equal(17D, explicitWords.Single(word => word.Text == "Heading").Letters[0].PointSize, 2);
        Assert.Equal(9D, explicitWords.Single(word => word.Text == "Bodyone").Letters[0].PointSize, 2);
        Assert.InRange(explicitWords.Single(word => word.Text == "Bodyone").Letters[0].StartBaseLine.X, 82.9, 83.1);
        Assert.Contains("Helvetica", explicitWords.Single(word => word.Text == "Deep").Letters[0].FontName, StringComparison.OrdinalIgnoreCase);

        options.Style = MarkdownPdfStyle.Plain();
        options.PdfOptions = null;
        using PdfCore.PdfDocument explicitStyle = source.ToPdf(options);
        using var explicitStyleArtifact = PdfPigDocument.Open(explicitStyle.ToBytes());
        Assert.Equal(11D, explicitStyleArtifact.GetPage(1).GetWords().Single(word => word.Text == "Bodyone").Letters[0].PointSize, 2);
        Assert.Null(options.PdfOptions);
    }
}
