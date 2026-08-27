using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlTextFormattingConversionTests {
    [Fact]
    public void ManagedHtmlRenderingComposesNestedScriptsAndSmallCapsAfterTextTransform() {
        const string html = """
            <p><sup><strong>Nested</strong></sup>
            <span style="font-variant:small-caps;text-transform:lowercase">MiXeD</span></p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions());
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText nested = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "Nested");
        HtmlRenderText smallCaps = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "MIXED");

        Assert.Equal(OfficeTextBaseline.Superscript, nested.Baseline);
        Assert.True((nested.Font.Style & OfficeFontStyle.Bold) == OfficeFontStyle.Bold);
        Assert.Equal("MIXED", smallCaps.Text);
    }

    [Fact]
    public void ManagedHtmlRenderingPreservesDecorationPatternsAndScriptsAcrossDrawingSvgAndRasterFormats() {
        const string html = """
            <p style="font-family:'Aptos';font-size:20px;color:#336699;font-weight:700;font-style:italic">
              <span style="text-decoration-line:underline line-through;text-decoration-style:wavy">Styled</span>
              <sup style="text-decoration-line:underline;text-decoration-style:double">Super</sup>
              <sub style="text-decoration-line:line-through;text-decoration-style:dotted">Sub</sub>
              <span style="font-variant:small-caps">SmallCaps</span>
            </p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, new HtmlRenderOptions());
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText styled = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "Styled");
        HtmlRenderText superscript = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "Super");
        HtmlRenderText subscript = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "Sub");
        HtmlRenderText smallCaps = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "SMALLCAPS");

        Assert.Equal(OfficeTextDecorationStyle.Wavy, styled.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Wavy, styled.StrikethroughStyle);
        Assert.Equal(OfficeTextBaseline.Superscript, superscript.Baseline);
        Assert.Equal(OfficeTextDecorationStyle.Double, superscript.UnderlineStyle);
        Assert.Equal(OfficeTextBaseline.Subscript, subscript.Baseline);
        Assert.Equal(OfficeTextDecorationStyle.Dotted, subscript.StrikethroughStyle);
        Assert.Equal("SMALLCAPS", smallCaps.Text);
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.FontVariantApproximated
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);

        OfficeDrawing drawing = page.CreateDrawing();
        OfficeDrawingText drawingSuper = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(), item => item.Text == "Super");
        Assert.Equal(OfficeTextBaseline.Superscript, drawingSuper.Baseline);
        Assert.Equal(OfficeTextDecorationStyle.Double, drawingSuper.UnderlineStyle);

        string svg = Encoding.UTF8.GetString(document.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains("text-decoration-style=\"wavy\"", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("text-decoration-style=\"double\"", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Styled", svg, StringComparison.Ordinal);

        foreach (OfficeImageExportFormat format in new[] {
                     OfficeImageExportFormat.Png,
                     OfficeImageExportFormat.Jpeg,
                     OfficeImageExportFormat.Tiff,
                     OfficeImageExportFormat.Webp
                 }) {
            OfficeImageExportResult image = document.ExportImage(format);
            Assert.Equal(format, image.Format);
            Assert.True(image.Bytes.Length > 32);
        }
    }

    [Fact]
    public void WordHtmlRoundTripRetainsNativeUnderlineDoubleStrikeBaselineCapsAndFontProperties() {
        using WordDocument source = WordDocument.Create();
        WordParagraph authored = source.AddParagraph("Styled");
        authored.SetBold().SetItalic().SetUnderline(WordUnderlineStyle.WavyDouble)
            .SetDoubleStrike().SetSuperScript().SetSmallCaps()
            .SetFontFamily("Aptos").SetFontSize(14).SetColorHex("336699");

        HtmlTextConversionResult export = source.ToHtmlResult(new WordToHtmlOptions { IncludeFontStyles = true });
        Assert.Contains("data-officeimo-word-underline=\"WavyDouble\"", export.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-word-double-strike=\"true\"", export.Value, StringComparison.Ordinal);
        Assert.Contains("<sup", export.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("font-variant:small-caps", export.Value, StringComparison.OrdinalIgnoreCase);

        using WordDocument imported = HtmlConversionDocument
            .Parse(export.Value, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToWordDocumentResult()
            .RequireValue();
        WordParagraph actual = Assert.Single(imported.Paragraphs);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(WordUnderlineStyle.WavyDouble, actual.Underline);
        Assert.True(actual.DoubleStrike);
        Assert.Equal(WordVerticalTextPosition.Superscript, actual.VerticalTextAlignment);
        Assert.Equal(WordCapsStyle.SmallCaps, actual.CapsStyle);
        Assert.Equal("Aptos", actual.FontFamily);
        Assert.Equal("336699", actual.ColorHex);
    }

    [Fact]
    public void ExcelSemanticHtmlRoundTripRetainsCellAndRichRunTypography() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Text");
        sheet.CellAt(1, 1).SetValue("Cell").SetBold().SetItalic()
            .SetUnderline(ExcelUnderlineStyle.DoubleAccounting).SetStrikethrough()
            .SetSuperscript().SetFontName("Aptos").SetFontSize(14).SetFontColor("336699");
        sheet.CellAt(2, 1).SetRichText(
            new ExcelRichTextRun("Rich") {
                Bold = true,
                Underline = true,
                UnderlineStyle = ExcelUnderlineStyle.SingleAccounting,
                FontName = "Aptos",
                FontSize = 13,
                FontColor = "663399"
            },
            new ExcelRichTextRun("Sub") {
                Italic = true,
                Strikethrough = true,
                VerticalTextAlignment = ExcelVerticalTextAlignment.Subscript
            });

        string html = source.ToHtml(ExcelHtmlSaveOptions.CreateSemanticTablesProfile());
        Assert.Contains("data-officeimo-excel-underline=\"DoubleAccounting\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-excel-underline=\"SingleAccounting\"", html, StringComparison.Ordinal);

        using ExcelDocument imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult()
            .RequireValue();
        ExcelSheet actualSheet = Assert.Single(imported.Sheets);
        ExcelCellStyleSnapshot cellStyle = actualSheet.GetCellStyle(1, 1);
        Assert.True(cellStyle.Bold);
        Assert.True(cellStyle.Italic);
        Assert.Equal(ExcelUnderlineStyle.DoubleAccounting, cellStyle.UnderlineStyle);
        Assert.True(cellStyle.Strikethrough);
        Assert.Equal(ExcelVerticalTextAlignment.Superscript, cellStyle.VerticalTextAlignment);
        Assert.Equal("Aptos", cellStyle.FontName);
        Assert.Equal(14D, cellStyle.FontSize);
        Assert.Equal("336699", cellStyle.FontColorHex);

        ExcelRichTextRun[] runs = actualSheet.GetRichText(2, 1).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Equal(ExcelUnderlineStyle.SingleAccounting, runs[0].UnderlineStyle);
        Assert.Equal("Aptos", runs[0].FontName);
        Assert.Equal(13D, runs[0].FontSize);
        Assert.Equal(ExcelVerticalTextAlignment.Subscript, runs[1].VerticalTextAlignment);
        Assert.True(runs[1].Strikethrough);
    }

    [Fact]
    public void ExcelSemanticHtmlRoundTripRetainsTypographyOnEmptyTemplateCells() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Template");
        sheet.CellAt(1, 1).SetValue("Anchor");
        sheet.CellAt(1, 2).SetBold().SetItalic()
            .SetUnderline(ExcelUnderlineStyle.DoubleAccounting).SetStrikethrough()
            .SetSubscript().SetFontName("Aptos").SetFontSize(14).SetFontColor("336699");

        string html = source.ToHtml(ExcelHtmlSaveOptions.CreateSemanticTablesProfile());
        Assert.Contains("data-officeimo-empty=\"true\"", html, StringComparison.Ordinal);

        using ExcelDocument imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult()
            .RequireValue();
        ExcelSheet actualSheet = Assert.Single(imported.Sheets);
        ExcelCellStyleSnapshot style = actualSheet.GetCellStyle(1, 2);

        Assert.False(actualSheet.TryGetCellValueSnapshot(1, 2, out _));
        Assert.True(style.Bold);
        Assert.True(style.Italic);
        Assert.Equal(ExcelUnderlineStyle.DoubleAccounting, style.UnderlineStyle);
        Assert.True(style.Strikethrough);
        Assert.Equal(ExcelVerticalTextAlignment.Subscript, style.VerticalTextAlignment);
        Assert.Equal("Aptos", style.FontName);
        Assert.Equal(14D, style.FontSize);
        Assert.Equal("336699", style.FontColorHex);
    }

    [Fact]
    public void ExcelSemanticHtmlRoundTripAppliesCellTypographyToRichRuns() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelCell cell = source.AddWorksheet("Text").CellAt(1, 1)
            .SetRichText(
                new ExcelRichTextRun("Inherited"),
                new ExcelRichTextRun(" Override") { Italic = true, FontColor = "CC0000" })
            .SetBold()
            .SetUnderline(ExcelUnderlineStyle.Double)
            .SetStrikethrough()
            .SetSuperscript()
            .SetFontName("Aptos")
            .SetFontSize(15)
            .SetFontColor("336699");

        string html = source.ToHtml(ExcelHtmlSaveOptions.CreateSemanticTablesProfile());
        using ExcelDocument imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult()
            .RequireValue();

        ExcelRichTextRun[] runs = imported.Sheets.Single().GetRichText(1, 1).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.All(runs, run => {
            Assert.True(run.Bold);
            Assert.Equal(ExcelUnderlineStyle.Double, run.UnderlineStyle);
            Assert.True(run.Strikethrough);
            Assert.Equal(ExcelVerticalTextAlignment.Superscript, run.VerticalTextAlignment);
            Assert.Equal("Aptos", run.FontName);
            Assert.Equal(15D, run.FontSize);
        });
        Assert.Equal("FF336699", runs[0].FontColor);
        Assert.Equal("FFCC0000", runs[1].FontColor);
        Assert.True(runs[1].Italic);
    }

    [Fact]
    public void PowerPointSemanticHtmlRoundTripRetainsNativeRunTypography() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointTextRun authored = source.AddSlide().AddTextBox("Styled")
            .Paragraphs.Single().Runs.Single();
        authored.Bold = true;
        authored.Italic = true;
        authored.UnderlineStyle = PowerPointUnderlineStyle.WavyDouble;
        authored.StrikeStyle = PowerPointStrikeStyle.Double;
        authored.Capitalization = PowerPointCapitalization.SmallCaps;
        authored.BaselinePercent = 37.5D;
        authored.FontName = "Aptos";
        authored.FontSizePoints = 14.5D;
        authored.Color = "336699";

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());
        Assert.Contains("data-officeimo-powerpoint-underline=\"WavyDouble\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-strike=\"Double\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-powerpoint-capitalization=\"SmallCaps\"", html, StringComparison.Ordinal);

        HtmlConversionDocument prepared = HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile());
        HtmlSemanticRun semanticRun = Assert.Single(prepared.CreateSemanticDocumentForConversion(HtmlCssMediaContext.Screen)
            .Sections.SelectMany(section => section.Blocks).SelectMany(block => block.Runs), run => run.Text == "Styled");
        Assert.True(semanticRun.Bold);
        Assert.True(semanticRun.Italic);
        Assert.Equal(OfficeTextDecorationStyle.Double, semanticRun.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, semanticRun.StrikethroughStyle);

        using PowerPointPresentation imported = prepared
            .ToPowerPointPresentationResult()
            .RequireValue();
        PowerPointTextRun actual = Assert.Single(imported.Slides).TextBoxes.Single()
            .Paragraphs.Single().Runs.Single();
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(PowerPointUnderlineStyle.WavyDouble, actual.UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, actual.StrikeStyle);
        Assert.Equal(PowerPointCapitalization.SmallCaps, actual.Capitalization);
        Assert.Equal(37.5D, actual.BaselinePercent);
        Assert.Equal("Aptos", actual.FontName);
        Assert.Equal(14.5D, actual.FontSizePoints);
        Assert.Equal("336699", actual.Color);
    }

    [Fact]
    public void PowerPointSemanticHtmlRoundTripRecreatesDynamicFieldsInAuthoredOrder() {
        using PowerPointPresentation source = PowerPointPresentation.Create();
        PowerPointParagraph paragraph = source.AddSlide().AddTextBox("Before ").Paragraphs.Single();
        paragraph.AddField("1", "slidenum", "{11111111-1111-1111-1111-111111111111}");
        paragraph.AddRun(" after");

        string html = source.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());
        Assert.Contains("data-officeimo-powerpoint-field=\"true\"", html, StringComparison.Ordinal);

        using PowerPointPresentation imported = HtmlConversionDocument
            .Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult()
            .RequireValue();
        IReadOnlyList<PowerPointParagraphInline> nodes = imported.Slides.Single().TextBoxes.Single()
            .Paragraphs.Single().InlineNodes;

        Assert.Equal(new[] { "Before ", "1", " after" }, nodes.Select(node => node.Text));
        PowerPointParagraphInline field = Assert.Single(nodes, node => node.Kind == PowerPointParagraphInlineKind.Field);
        Assert.Equal("slidenum", field.FieldType);
        Assert.Equal("{11111111-1111-1111-1111-111111111111}", field.FieldId);
    }

    [Fact]
    public void InvalidNumericNativeStyleMetadataFallsBackWithoutThrowingAcrossOfficeTargets() {
        using WordDocument word = HtmlConversionDocument.Parse(
            "<p><span data-officeimo-word-underline=\"999\">Word</span></p>")
            .ToWordDocumentResult().RequireValue();
        Assert.Null(Assert.Single(word.Paragraphs).Underline);

        using ExcelDocument excelSource = ExcelDocument.Create();
        excelSource.AddWorksheet("Text").CellAt(1, 1).SetRichText(
            new ExcelRichTextRun("Ex") { Bold = true },
            new ExcelRichTextRun("cel") { Italic = true });
        string excelHtml = excelSource.ToHtml(ExcelHtmlSaveOptions.CreateSemanticTablesProfile())
            .Replace("<span", "<span data-officeimo-excel-underline=\"999\"");
        using ExcelDocument excel = HtmlConversionDocument.Parse(excelHtml, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToExcelDocumentResult().RequireValue();
        Assert.All(Assert.Single(excel.Sheets).GetRichText(1, 1), run => Assert.Null(run.UnderlineStyle));

        using PowerPointPresentation powerPointSource = PowerPointPresentation.Create();
        powerPointSource.AddSlide().AddTextBox("PowerPoint");
        string powerPointHtml = powerPointSource.ToHtml(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile())
            .Replace("<span", "<span data-officeimo-powerpoint-underline=\"999\" data-officeimo-powerpoint-strike=\"999\" data-officeimo-powerpoint-capitalization=\"999\"");
        using PowerPointPresentation powerPoint = HtmlConversionDocument.Parse(
            powerPointHtml, HtmlConversionDocumentOptions.CreateTrustedProfile())
            .ToPowerPointPresentationResult().RequireValue();
        PowerPointTextRun run = powerPoint.Slides.Single().TextBoxes.Single().Paragraphs.Single().Runs.Single();
        Assert.Null(run.UnderlineStyle);
        Assert.Null(run.StrikeStyle);
        Assert.Null(run.Capitalization);
    }

    [Fact]
    public void HtmlToOneNoteRetainsRepresentableRunTypography() {
        const string html = """
            <p><span style="font-family:'Aptos';font-size:14pt;color:#336699;background-color:#FFF2CC;"
                ><strong><em><u><s><sup>Styled</sup></s></u></em></strong></span></p>
            """;

        OneNoteSection section = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult().RequireValue();
        OneNoteTextRun run = Assert.Single(Assert.Single(Assert.Single(section.Pages).Outlines).Children
            .OfType<OneNoteParagraph>().Single().Runs);

        Assert.True(run.Style.Bold);
        Assert.True(run.Style.Italic);
        Assert.True(run.Style.Underline);
        Assert.True(run.Style.Strikethrough);
        Assert.True(run.Style.Superscript);
        Assert.NotEqual(true, run.Style.Subscript);
        Assert.Equal("Aptos", run.Style.FontFamily);
        Assert.Equal(14D, run.Style.FontSize);
        Assert.Equal(0xFF336699U, run.Style.ColorArgb);
        Assert.Equal(0xFFFFF2CCU, run.Style.HighlightColorArgb);
    }

    [Fact]
    public void HtmlBlockBackgroundDoesNotBecomeOneNoteRunHighlightButInlineBackgroundDoes() {
        const string html = "<p style=\"background-color:#112233\">Plain <span style=\"background-color:#FFF2CC\">Marked</span></p>";

        OneNoteSection section = HtmlConversionDocument.Parse(html).ToOneNoteSectionResult().RequireValue();
        OneNoteTextRun[] runs = Assert.Single(Assert.Single(Assert.Single(section.Pages).Outlines).Children
            .OfType<OneNoteParagraph>()).Runs.ToArray();

        Assert.Null(Assert.Single(runs, run => run.Text == "Plain ").Style.HighlightColorArgb);
        Assert.Equal(0xFFFFF2CCU, Assert.Single(runs, run => run.Text == "Marked").Style.HighlightColorArgb);
    }

    [Fact]
    public void OneNoteSemanticHtmlRoundTripRetainsNativeRunTypography() {
        var source = new OneNoteSection { Name = "Text" };
        var page = new OneNotePage { Title = "Styled" };
        var paragraph = new OneNoteParagraph();
        var authored = new OneNoteTextRun { Text = "Styled" };
        authored.Style.FontFamily = "Aptos";
        authored.Style.FontSize = 14D;
        authored.Style.ColorArgb = 0xFF336699U;
        authored.Style.HighlightColorArgb = 0xFFFFF2CCU;
        authored.Style.Bold = true;
        authored.Style.Italic = true;
        authored.Style.Underline = true;
        authored.Style.Strikethrough = true;
        authored.Style.Subscript = true;
        paragraph.Runs.Add(authored);
        page.DirectContent.Add(paragraph);
        source.Pages.Add(page);

        HtmlTextConversionResult exported = source.ToHtmlDocumentResult();
        Assert.Contains("font-family:&quot;Aptos&quot;", exported.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("font-size:14pt", exported.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("color:#336699", exported.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("background-color:#FFF2CC", exported.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<u><s>", exported.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<sub", exported.Value, StringComparison.OrdinalIgnoreCase);

        OneNoteSection imported = HtmlConversionDocument.Parse(exported.Value).ToOneNoteSectionResult().RequireValue();
        OneNoteTextRun actual = Assert.Single(imported.Pages.SelectMany(importedPage => importedPage.Outlines)
            .SelectMany(outline => outline.Children).OfType<OneNoteParagraph>()
            .SelectMany(item => item.Runs), item => item.Text == "Styled" && item.Style.FontFamily == "Aptos");
        Assert.Equal("Aptos", actual.Style.FontFamily);
        Assert.Equal(14D, actual.Style.FontSize);
        Assert.Equal(0xFF336699U, actual.Style.ColorArgb);
        Assert.Equal(0xFFFFF2CCU, actual.Style.HighlightColorArgb);
        Assert.True(actual.Style.Bold);
        Assert.True(actual.Style.Italic);
        Assert.True(actual.Style.Underline);
        Assert.True(actual.Style.Strikethrough);
        Assert.True(actual.Style.Subscript);
    }

    [Fact]
    public void StyledHtmlToPdfToEveryImageFormatKeepsAVisualArtifact() {
        const string html = """
            <p style="font-family:'Aptos';font-size:24px;color:#336699;font-weight:700;font-style:italic">
              <span style="text-decoration-line:underline line-through;text-decoration-style:wavy">Styled</span>
              <sup>Super</sup><sub>Sub</sub>
            </p>
            """;
        PdfDocumentConversionResult pdf = HtmlConversionDocument.Parse(html).ToPdfDocumentResult();
        Assert.True(pdf.ToBytes().Length > 100);

        foreach (OfficeImageExportFormat format in new[] {
                     OfficeImageExportFormat.Png,
                     OfficeImageExportFormat.Svg,
                     OfficeImageExportFormat.Jpeg,
                     OfficeImageExportFormat.Tiff,
                     OfficeImageExportFormat.Webp
                 }) {
            OfficeImageExportResult image = Assert.Single(pdf.ExportImages(format));
            Assert.Equal(format, image.Format);
            Assert.True(image.Bytes.Length > 32);
        }
    }

    [Fact]
    public void StyledWordDocumentReachesEveryDirectAndPdfBackedImageFormat() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Styled").SetBold().SetItalic()
            .SetUnderline(WordUnderlineStyle.WavyDouble).SetDoubleStrike()
            .SetSuperScript().SetFontFamily("Aptos").SetFontSize(18).SetColorHex("336699");
        OfficeImageExportFormat[] formats = {
            OfficeImageExportFormat.Png,
            OfficeImageExportFormat.Svg,
            OfficeImageExportFormat.Jpeg,
            OfficeImageExportFormat.Tiff,
            OfficeImageExportFormat.Webp
        };

        foreach (OfficeImageExportFormat format in formats) {
            OfficeImageExportResult direct = Assert.Single(document.ExportImages(format));
            Assert.True(direct.Bytes.Length > 32);
        }

        PdfDocumentConversionResult pdf = document.ToPdfDocumentResult();
        foreach (OfficeImageExportFormat format in formats) {
            OfficeImageExportResult flattened = Assert.Single(pdf.ExportImages(format));
            Assert.True(flattened.Bytes.Length > 32);
        }
    }

    [Fact]
    public void ConversionCatalogAdvertisesEverySharedImageFormatForEveryImageSource() {
        string[] sources = { "docx", "xlsx", "pptx", "html", "onenote", "visio", "email", "epub", "odt", "ods", "odp", "pdf" };
        string[] formats = { "png", "svg", "jpeg", "tiff", "webp" };

        foreach (string source in sources) {
            foreach (string format in formats) {
                OfficeConversionCapability? route = OfficeConversionCapabilityCatalog.Find(source + "-" + format);
                Assert.NotNull(route);
                Assert.Equal(OfficeConversionFidelityKind.FixedLayout, route!.Fidelity);
                Assert.Equal("IReadOnlyList<OfficeImageExportResult>", route.ResultContract);
                Assert.Equal(
                    format == "svg"
                        ? OfficeConversionTextFormattingKind.VectorAppearance
                        : OfficeConversionTextFormattingKind.FlattenedRaster,
                    route.TextFormatting);
                Assert.False(string.IsNullOrWhiteSpace(route.TextFormattingContract));
            }
        }
    }

    [Fact]
    public void EveryAdvertisedConversionDeclaresItsTextFormattingContract() {
        Assert.NotEmpty(OfficeConversionCapabilityCatalog.All);
        Assert.All(OfficeConversionCapabilityCatalog.All, route => {
            Assert.NotEqual(OfficeConversionTextFormattingKind.Unspecified, route.TextFormatting);
            Assert.False(string.IsNullOrWhiteSpace(route.TextFormattingContract));
            if (route.Source == "PDF" && route.Target is not ("PNG" or "SVG" or "JPEG" or "TIFF" or "WebP")) {
                Assert.Equal(OfficeConversionTextFormattingKind.ReconstructedFromFixedLayout, route.TextFormatting);
            }
            if (route.Target == "PDF") {
                Assert.Equal(OfficeConversionTextFormattingKind.FixedLayoutAppearance, route.TextFormatting);
            }
        });
    }

    [Fact]
    public void AdjacentDocumentConvertersDeclareTheirActualFormattingBoundary() {
        string[] syntaxSubsetRoutes = {
            "adf-markdown", "markdown-adf", "adf-html", "html-adf",
            "markdown-confluence", "html-confluence", "confluence-markdown", "confluence-html",
            "officemarkup-docx", "officemarkup-xlsx", "officemarkup-pptx"
        };

        Assert.All(syntaxSubsetRoutes, id => {
            OfficeConversionCapability route = Assert.IsType<OfficeConversionCapability>(
                OfficeConversionCapabilityCatalog.Find(id));
            Assert.Equal(OfficeConversionTextFormattingKind.SyntaxSubset, route.TextFormatting);
        });

        Assert.Equal(OfficeConversionInputKind.ObjectModel,
            OfficeConversionCapabilityCatalog.Find("confluence-markdown")!.InputKind);
        Assert.Equal(OfficeConversionInputKind.ObjectModel,
            OfficeConversionCapabilityCatalog.Find("confluence-html")!.InputKind);

        Assert.Equal(OfficeConversionTextFormattingKind.DataOnly,
            OfficeConversionCapabilityCatalog.Find("csv-xlsx")!.TextFormatting);
        Assert.Equal(OfficeConversionTextFormattingKind.DataOnly,
            OfficeConversionCapabilityCatalog.Find("xlsx-csv")!.TextFormatting);
    }
}
