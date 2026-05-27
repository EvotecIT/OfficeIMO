using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class PdfComposePageOptionsTests {
        [Fact]
        public void ComposePage_RejectsNullConfigurationDelegates() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultTextStyle((Action<PdfTextStyleCompose>)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultTextStyle((PdfTextStyle)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Theme(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultParagraphStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultTableStyle((PdfTableStyle)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultTableStyle((string)null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultHeadingStyle(1, null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultListStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultPanelStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultHorizontalRuleStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultImageStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultDrawingStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultRowStyle(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(content => content.Item(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(content => content.Column(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(content => content.Column(column => column.Item(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(content => content.Column(column => column.Item().Element(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Content(content => content.Row(null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(null!))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Header(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Footer(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Page(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Section(null!));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Section(null!)));
        }

        [Fact]
        public void ComposePage_RejectsInvalidDefaultTextStyleFont() {
            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.DefaultTextStyle(style => style.Font((PdfStandardFont)99)))));

            Assert.Equal("font", exception.ParamName);
            Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposeContent_ItemAndSpacerProvideDirectWordLikeFlow() {
            byte[] pdfBytes = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Item(item => item
                            .H1("DirectComposeTitle")
                            .Paragraph(paragraph => paragraph.Text("DirectComposeLead")))
                        .Spacer(24)
                        .Column(column => column
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnComposeTop")))
                            .Spacer(18)
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnComposeBottom")))))))
                .ToBytes();

            string text = PdfReadDocument.Load(pdfBytes).ExtractText();
            Assert.Contains("DirectComposeTitle", text, StringComparison.Ordinal);
            Assert.Contains("DirectComposeLead", text, StringComparison.Ordinal);
            Assert.Contains("ColumnComposeTop", text, StringComparison.Ordinal);
            Assert.Contains("ColumnComposeBottom", text, StringComparison.Ordinal);

            using var pdf = PdfDocument.Open(new MemoryStream(pdfBytes));
            var page = pdf.GetPage(1);
            double leadY = FindWordStartY(page, "DirectComposeLead");
            double columnTopY = FindWordStartY(page, "ColumnComposeTop");
            double columnBottomY = FindWordStartY(page, "ColumnComposeBottom");

            Assert.True(leadY - columnTopY >= 32, $"Expected direct content spacer to preserve visible rhythm. Lead y: {leadY:0.##}, top y: {columnTopY:0.##}.");
            Assert.True(columnTopY - columnBottomY >= 26, $"Expected column spacer to preserve visible rhythm. Top y: {columnTopY:0.##}, bottom y: {columnBottomY:0.##}.");
        }

        [Fact]
        public void ComposeContent_PageBreaksProvideDirectWordLikeFlow() {
            byte[] pdfBytes = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Item(item => item.Paragraph(paragraph => paragraph.Text("DirectPageOne")))
                        .PageBreak()
                        .Column(column => column
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnPageTwo")))
                            .PageBreak()
                            .Item(item => item.Paragraph(paragraph => paragraph.Text("ColumnPageThree")))))))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(pdfBytes));
            Assert.Equal(3, pdf.NumberOfPages);
            Assert.Contains("DirectPageOne", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("ColumnPageTwo", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("ColumnPageTwo", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("ColumnPageThree", pdf.GetPage(2).Text, StringComparison.Ordinal);
            Assert.Contains("ColumnPageThree", pdf.GetPage(3).Text, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposeItem_ElementPageBreakProvidesNestedWordLikeFlow() {
            byte[] pdfBytes = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 12
                })
                .Compose(compose => compose.Page(page => page
                    .Margin(72)
                    .Content(content => content
                        .Column(column => column
                            .Item(item => item
                                .Paragraph(paragraph => paragraph.Text("NestedPageOne"))
                                .Element(element => element
                                    .PageBreak()
                                    .Paragraph(paragraph => paragraph.Text("NestedPageTwo"))))))))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(pdfBytes));
            Assert.Equal(2, pdf.NumberOfPages);
            Assert.Contains("NestedPageOne", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain("NestedPageTwo", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("NestedPageTwo", pdf.GetPage(2).Text, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_RejectsInvalidPageSetupScalarsAtAssignment() {
            var pageWidthException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Size(0, 792))));
            Assert.Equal("width", pageWidthException.ParamName);

            var pageHeightException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Size(612, double.NaN))));
            Assert.Equal("height", pageHeightException.ParamName);

            var pageSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PageSize(612, double.PositiveInfinity));
            Assert.Equal("height", pageSizeException.ParamName);

            var defaultPageSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Size(default))));
            Assert.Equal("size", defaultPageSizeException.ParamName);

            var uniformMarginException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Margin(-1))));
            Assert.Equal("all", uniformMarginException.ParamName);

            var sideMarginException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Margin(10, 20, double.NegativeInfinity, 20))));
            Assert.Equal("right", sideMarginException.ParamName);

            var pageMarginsException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PageMargins(10, double.NaN, 10, 10));
            Assert.Equal("top", pageMarginsException.ParamName);

            var documentPageNumberException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().PageNumberStart(0));
            Assert.Equal("PageNumberStart", documentPageNumberException.ParamName);

            var sectionPageNumberException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Section(section => section.PageNumberStart(0)));
            Assert.Equal("PageNumberStart", sectionPageNumberException.ParamName);

            var pageNumberStyleException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().PageNumberStyle((PdfPageNumberStyle)99));
            Assert.Equal("PageNumberStyle", pageNumberStyleException.ParamName);
        }

        [Fact]
        public void ComposePage_AllowsMarginsBeforeLargerPageSizeAndKeepsImpossibleFrameRenderTime() {
            var doc = PdfDoc.Create();
            doc.Compose(c => c.Page(page => {
                page.Margin(400);
                page.Size(1000, 1000);
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Large page after margins."))));
            }));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            string text = Normalize(pdf.GetPage(1).Text);
            Assert.Contains("Largepageaftermargins", text, StringComparison.OrdinalIgnoreCase);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => {
                    page.Size(200, 200);
                    page.Margin(100);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Impossible frame."))));
                })).ToBytes());

            Assert.Contains("PDF margins must leave a positive content width.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void PageMarginsPresets_MatchWordCompatiblePointValues() {
            AssertMargins(PageMargins.Normal, 72, 72, 72, 72);
            AssertMargins(PageMargins.Narrow, 36, 36, 36, 36);
            AssertMargins(PageMargins.Moderate, 54, 72, 54, 72);
            AssertMargins(PageMargins.Wide, 144, 72, 144, 72);
            AssertMargins(PageMargins.Mirrored, 90, 72, 72, 72);
            AssertMargins(PageMargins.Office2003Default, 90, 72, 90, 72);
            AssertMargins(PageMargins.Uniform(24), 24, 24, 24, 24);
        }

        [Fact]
        public void PageSetupValues_CanBeCreatedFromOfficeFriendlyUnits() {
            var letter = PageSize.FromInches(8.5, 11);
            Assert.Equal(612, letter.Width);
            Assert.Equal(792, letter.Height);

            var a4 = PageSize.FromCentimeters(21, 29.7);
            Assert.InRange(a4.Width, 595.2, 595.4);
            Assert.InRange(a4.Height, 841.8, 842.0);

            AssertMargins(PageMargins.UniformInches(0.5), 36, 36, 36, 36);

            var customInches = PageMargins.FromInches(1, 1.25, 1.5, 2);
            AssertMargins(customInches, 72, 90, 108, 144);

            var customCentimeters = PageMargins.FromCentimeters(2.54, 1.27, 3.81, 5.08);
            AssertMargins(customCentimeters, 72, 36, 108, 144);

            var sizeException = Assert.Throws<ArgumentOutOfRangeException>(() => PageSize.FromInches(0, 11));
            Assert.Equal("width", sizeException.ParamName);

            var marginException = Assert.Throws<ArgumentOutOfRangeException>(() => PageMargins.UniformCentimeters(double.NaN));
            Assert.Equal("centimeters", marginException.ParamName);
        }

        [Fact]
        public void PdfOptionsPageSetupProperties_ApplyReusableValuesAndValidate() {
            var options = new PdfOptions {
                PageSize = PageSizes.A4.Landscape(),
                Margins = PageMargins.Moderate
            };

            Assert.InRange(options.PageWidth, 841.0, 843.0);
            Assert.InRange(options.PageHeight, 594.0, 596.0);
            Assert.Equal(PdfPageOrientation.Landscape, options.PageOrientation);
            AssertMargins(options.Margins, 54, 72, 54, 72);

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PdfOptions { PageSize = default });

            Assert.Equal("PageSize", exception.ParamName);
        }

        [Fact]
        public void PdfDocPageSetupFluent_AppliesToTopLevelFlowAndComposePages() {
            var doc = PdfDoc.Create()
                .Size(PageSizes.A5)
                .Landscape()
                .Margin(PageMargins.Narrow)
                .Paragraph(p => p.Text("Document default page setup."));

            doc.Compose(c => c.Page(page =>
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Composed page inherits setup."))))));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);

            Assert.InRange(page1.Width, 594.0, 596.0);
            Assert.InRange(page1.Height, 419.0, 421.0);
            Assert.InRange(page2.Width, 594.0, 596.0);
            Assert.InRange(page2.Height, 419.0, 421.0);
            Assert.InRange(FindWordStartX(page1, "Document"), 35.5, 36.5);
            Assert.InRange(FindWordStartX(page2, "Composed"), 35.5, 36.5);
            Assert.Contains("Documentdefaultpagesetup", Normalize(page1.Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Composedpageinheritssetup", Normalize(page2.Text), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ComposePageMarginPreset_AppliesReusableMarginValue() {
            var doc = PdfDoc.Create();

            doc.Compose(c => c.Page(page => {
                page.Size(PageSizes.A4);
                page.Margin(PageMargins.Narrow);
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Narrow margin body."))));
            }));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            var firstBodyLetter = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .OrderByDescending(letter => letter.StartBaseLine.Y)
                .ThenBy(letter => letter.StartBaseLine.X)
                .First();

            Assert.InRange(firstBodyLetter.StartBaseLine.X, 35.5, 36.5);
            Assert.Contains("Narrowmarginbody", Normalize(pdf.GetPage(1).Text), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterCompose_RejectsInvalidTypographyAndPlacementValues() {
            var headerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Header(header => header.Font((PdfStandardFont)99)))));

            Assert.Equal("HeaderFont", headerFontException.ParamName);
            Assert.Contains("PDF header font must be one of the supported standard PDF fonts.", headerFontException.Message, StringComparison.Ordinal);

            var footerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.Font((PdfStandardFont)99)))));

            Assert.Equal("FooterFont", footerFontException.ParamName);
            Assert.Contains("PDF footer font must be one of the supported standard PDF fonts.", footerFontException.Message, StringComparison.Ordinal);

            var headerSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Header(header => header.FontSize(double.NaN)))));

            Assert.Equal("size", headerSizeException.ParamName);

            var footerSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.FontSize(0)))));

            Assert.Equal("size", footerSizeException.ParamName);

            var headerOffsetException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Header(header => header.Offset(double.NegativeInfinity)))));

            Assert.Equal("points", headerOffsetException.ParamName);

            var footerOffsetException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page =>
                    page.Footer(footer => footer.Offset(-1)))));

            Assert.Equal("points", footerOffsetException.ParamName);

            var renderHeaderOffsetException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => {
                    page.Margin(20);
                    page.Header(header => header.Offset(21).Text("Invalid header offset"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Body"))));
                })).ToBytes());

            Assert.Contains("PDF header offset must not exceed the top margin when header content is enabled.", renderHeaderOffsetException.Message, StringComparison.Ordinal);

            var renderFooterOffsetException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => {
                    page.Margin(20);
                    page.Footer(footer => footer.Offset(21).PageNumber());
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Body"))));
                })).ToBytes());

            Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", renderFooterOffsetException.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void StandardFontNames_RejectInvalidValuesInsteadOfFallingBack() {
            Assert.Equal("Helvetica", PdfStandardFont.Helvetica.ToBaseFontName());
            Assert.Equal("Times-Roman", PdfStandardFont.TimesRoman.ToBaseFontName());
            Assert.Equal("Courier-BoldOblique", PdfStandardFont.CourierBoldOblique.ToBaseFontName());

            var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
                ((PdfStandardFont)99).ToBaseFontName());

            Assert.Equal("font", exception.ParamName);
            Assert.Contains("PDF font must be one of the supported standard PDF fonts.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void StandardFontDictionaryBuilder_EmitsType1WinAnsiFontObjects() {
            Assert.Equal(
                "<< /Type /Font /Subtype /Type1 /BaseFont /Times-BoldItalic /Encoding /WinAnsiEncoding >>\n",
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject(PdfStandardFont.TimesBoldItalic));

            PdfDictionary dictionary = PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(PdfStandardFont.CourierOblique);

            Assert.Equal("Font", Assert.IsType<PdfName>(dictionary.Items["Type"]).Name);
            Assert.Equal("Type1", Assert.IsType<PdfName>(dictionary.Items["Subtype"]).Name);
            Assert.Equal("Courier-Oblique", Assert.IsType<PdfName>(dictionary.Items["BaseFont"]).Name);
            Assert.Equal("WinAnsiEncoding", Assert.IsType<PdfName>(dictionary.Items["Encoding"]).Name);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfStandardFontDictionaryBuilder.BuildStandardType1FontObject((PdfStandardFont)99));
        }

        [Fact]
        public void CatalogDictionaryBuilder_EmitsGeneratedCatalogsAndSharedEntries() {
            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 5));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Names << /Dests 7 0 R >> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 7));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines /Names << /Dests 7 0 R >> >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 5, 7));

            Assert.Equal(
                "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R >>\n",
                PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, 8));

            var sb = new StringBuilder();
            PdfCatalogDictionaryBuilder.AppendCatalogStart(sb, 3);
            PdfCatalogDictionaryBuilder.AppendNameEntry(sb, "PageLayout", "TwoColumnLeft");
            PdfCatalogDictionaryBuilder.AppendReferenceEntry(sb, "Outlines", 9);
            sb.Append(" >>\n");

            Assert.Equal("<< /Type /Catalog /Pages 3 0 R /PageLayout /TwoColumnLeft /Outlines 9 0 R >>\n", sb.ToString());
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(0, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfCatalogDictionaryBuilder.BuildGeneratedCatalogDictionary(2, 0, 0, -1));
        }

        [Fact]
        public void PageDictionaryBuilder_EmitsGeneratedPageWithResourcesAndAnnotations() {
            var fonts = new[] { ("F1", 4) };
            var xobjects = new[] { ("/Im1", 5) };
            var graphicsStates = new[] { ("/GS1", 6) };
            var shadings = new[] { ("/Sh1", 7) };
            var annotations = new[] { 8, 9 };

            string page = PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(
                2,
                612,
                792,
                10,
                fonts,
                xobjects,
                graphicsStates,
                shadings,
                annotations);

            Assert.Equal(
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject << /Im1 5 0 R >> /ExtGState << /GS1 6 0 R >> /Shading << /Sh1 7 0 R >> >> /Contents 10 0 R /Annots [ 8 0 R 9 0 R ] >>\n",
                page);

            Assert.Equal(
                " /XObject << /Image#201 3 0 R >>",
                PdfPageDictionaryBuilder.BuildResourcePart("XObject", new[] { ("Image 1", 3) }));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfPageDictionaryBuilder.BuildGeneratedPageDictionary(2, 0, 792, 10, fonts, xobjects, graphicsStates, shadings, annotations));
        }

        [Fact]
        public void AnnotationDictionaryBuilder_EmitsUriLinkAnnotationsWithEscapedUri() {
            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Rect [10 20.5 110 44.25] /A << /S /URI /URI (https://evotec.xyz/docs\\(pdf\\)) >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20.5, 110, 44.25, "https://evotec.xyz/docs(pdf)"));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Link /Border [0 0 0] /Contents (Jump metadata) /Rect [10 20.5 110 44.25] /A << /S /GoTo /D (Intro\\(A\\)) >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20.5, 110, 44.25, "Intro(A)", "Jump metadata"));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Person.Name) /V <416461> /DV <416461> /Rect [10 20.5 110 44.25] /F 4 /DA (/Helv 10 Tf 0 g) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildTextFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Person.Name", "Ada", 10, 12));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Btn /T (AcceptTerms) /V /Yes /DV /Yes /Rect [10 20.5 26 36.5] /F 4 /AS /Yes /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N << /Off 12 0 R /Yes 13 0 R >> >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20.5, 26, 36.5, "AcceptTerms", true, "Yes", 12, 13));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T (Country) /V <506F6C616E64> /DV <506F6C616E64> /Opt [ <506F6C616E64> <556E6974656420537461746573> ] /Ff 131072 /Rect [10 20.5 110 44.25] /F 4 /DA (/Helv 10 Tf 0 g) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 44.25, "Country", new[] { "Poland", "United States" }, "Poland", 10, 12, isComboBox: true));

            Assert.Equal(
                "<< /Type /Annot /Subtype /Widget /FT /Ch /T (Countries) /V [<506F6C616E64> <556E6974656420537461746573>] /DV [<506F6C616E64> <556E6974656420537461746573>] /Opt [ <506F6C616E64> <4765726D616E79> <556E6974656420537461746573> ] /Ff 2097152 /Rect [10 20.5 110 70] /F 4 /DA (/Helv 10 Tf 0 g) /MK << /BC [0.75 0.75 0.75] /BG [1 1 1] >> /AP << /N 12 0 R >> >>\n",
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20.5, 110, 70, "Countries", new[] { "Poland", "Germany", "United States" }, new[] { "Poland", "United States" }, 10, 12, isComboBox: false, allowsMultipleSelection: true));

            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, 44, "not-a-uri"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 10, 44, "https://evotec.xyz"));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfAnnotationDictionaryBuilder.BuildUriLinkAnnotation(10, 20, 110, double.NaN, "https://evotec.xyz"));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(10, 20, 110, 44, " "));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildCheckBoxWidgetAnnotation(10, 20, 26, 36, "AcceptTerms", true, "Off", 12, 13));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Country", Array.Empty<string>(), "Poland", 10, 12, isComboBox: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Country", new[] { "Poland" }, "Germany", 10, 12, isComboBox: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland" }, Array.Empty<string>(), 10, 12, isComboBox: false, allowsMultipleSelection: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland" }, new[] { "Poland", "Poland" }, 10, 12, isComboBox: false, allowsMultipleSelection: true));
            Assert.Throws<ArgumentException>(() =>
                PdfAnnotationDictionaryBuilder.BuildChoiceFieldWidgetAnnotation(10, 20, 110, 44, "Countries", new[] { "Poland", "Germany" }, new[] { "Poland", "Germany" }, 10, 12, isComboBox: true, allowsMultipleSelection: true));
        }

        [Fact]
        public void AcroFormDictionaryBuilder_EmitsFieldsAndTextAppearance() {
            Assert.Equal(
                "<< /Fields [ 4 0 R 5 0 R ] /NeedAppearances true /DR << /Font << /Helv 3 0 R >> >> /DA (/Helv 10 Tf 0 g) >>\n",
                PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(new[] { 4, 5 }, 3));

            string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(120, 20, "Ada", 10);
            Assert.Contains("<416461> Tj", content);

            string dictionary = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(120, 20, 3, Encoding.ASCII.GetByteCount(content));
            Assert.Equal(
                "<< /Type /XObject /Subtype /Form /BBox [0 0 120 20] /Resources << /Font << /Helv 3 0 R >> >> /Length " + Encoding.ASCII.GetByteCount(content) + " >>",
                dictionary);

            string checkedContent = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(16, 16, selected: true);
            Assert.Contains(" l S", checkedContent);

            string checkBoxDictionary = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(16, 16, Encoding.ASCII.GetByteCount(checkedContent));
            Assert.Equal(
                "<< /Type /XObject /Subtype /Form /BBox [0 0 16 16] /Length " + Encoding.ASCII.GetByteCount(checkedContent) + " >>",
                checkBoxDictionary);

            Assert.Throws<ArgumentException>(() => PdfAcroFormDictionaryBuilder.BuildAcroFormDictionary(Array.Empty<int>(), 3));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceStreamDictionary(120, 20, 3, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceStreamDictionary(16, 16, -1));
        }

        [Fact]
        public void OutlineDictionaryBuilder_EmitsRootAndNestedItems() {
            Assert.Equal(
                "<< /Type /Outlines /First 6 0 R /Last 9 0 R /Count 3 >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineRoot(6, 9, 3));

            Assert.Equal(
                "<< /Title (Chapter \\(A\\)) /Parent 5 0 R /Prev 6 0 R /Next 8 0 R /First 10 0 R /Last 11 0 R /Count 2 /Dest [3 0 R /XYZ 0 712.25 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Chapter (A)", 5, 6, 8, 10, 11, 2, 3, 712.25));

            Assert.Equal(
                "<< /Title (Leaf) /Parent 5 0 R /Dest [3 0 R /XYZ 0 0 0] >>\n",
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Leaf", 5, 0, 0, 0, 0, 0, 3, 0));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineRoot(6, 9, -1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Bad", 5, 0, 0, 0, 0, -1, 3, 10));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfOutlineDictionaryBuilder.BuildOutlineItem("Bad", 5, 0, 0, 0, 0, 0, 3, double.PositiveInfinity));
        }

        [Fact]
        public void VisualResourceDictionaryBuilder_EmitsOpacityAndAxialShadingResources() {
            Assert.Equal(
                "<< /Type /ExtGState /ca 0.35 /CA 0.75 >>\n",
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(0.35, 0.75));

            Assert.Equal(
                "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [30 118 120 118] /Function << /FunctionType 2 /Domain [0 1] /C0 [0.039 0.078 0.118] /C1 [1 0.502 0] /N 1 >> /Extend [true true] >>\n",
                PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(
                    30,
                    118,
                    120,
                    118,
                    OfficeColor.FromRgb(10, 20, 30),
                    OfficeColor.FromRgb(255, 128, 0)));

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(-0.1, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildExtGStateObject(1, 1.1));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(double.NaN, 0, 1, 1, OfficeColor.Black, OfficeColor.White));
        }

        [Fact]
        public void FooterText_RejectsNullConfigurationAndTextSegments() {
            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.Text((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText((Action<HeaderTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text((Action<FooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText((Action<FooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((string)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText((Action<FooterTextBuilder>)null!)))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Text(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.Text(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.EvenPagesText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.FirstPageText(text => text.Text(null!))))));

            Assert.Throws<ArgumentNullException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesText(text => text.Text(null!))))));

            Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.Zones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.Zones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Header(header => header.FirstPageZones(null, null, null)))));

            Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Footer(footer => footer.EvenPagesZones(null, null, null)))));
        }

        [Fact]
        public void FooterSegments_RejectInvalidExternalState() {
            var nullEntryOptions = new PdfOptions {
                ShowPageNumbers = true,
                FooterSegments = new System.Collections.Generic.List<FooterSegment> { null! }
            };

            var nullEntryException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create(nullEntryOptions)
                    .Paragraph(p => p.Text("Invalid footer segment"))
                    .ToBytes());
            Assert.Contains("footer segments", nullEntryException.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FooterSegment_RejectsInvalidIntrinsicStateAtConstruction() {
            var nullTextException = Assert.Throws<ArgumentNullException>(() =>
                new FooterSegment(FooterSegmentKind.Text, null));
            Assert.Equal("text", nullTextException.ParamName);
            Assert.Contains("Footer text segments cannot be null.", nullTextException.Message, StringComparison.Ordinal);

            var invalidKindException = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new FooterSegment((FooterSegmentKind)99));
            Assert.Equal("kind", invalidKindException.ParamName);
            Assert.Contains("Footer segments must use a supported segment kind.", invalidKindException.Message, StringComparison.Ordinal);

            var textSegment = new FooterSegment(FooterSegmentKind.Text, string.Empty);
            var pageSegment = new FooterSegment(FooterSegmentKind.PageNumber);
            var totalSegment = new FooterSegment(FooterSegmentKind.TotalPages);

            Assert.Equal(string.Empty, textSegment.Text);
            Assert.Null(pageSegment.Text);
            Assert.Null(totalSegment.Text);
        }

        [Fact]
        public void FooterSegments_SnapshotAssignedAndReadbackLists() {
            var assigned = new System.Collections.Generic.List<FooterSegment> {
                new FooterSegment(FooterSegmentKind.Text, "Page "),
                new FooterSegment(FooterSegmentKind.PageNumber)
            };

            var options = new PdfOptions {
                ShowPageNumbers = true,
                FooterSegments = assigned
            };

            assigned[0] = new FooterSegment(FooterSegmentKind.Text, "Mutated");
            assigned.Add(new FooterSegment(FooterSegmentKind.TotalPages));

            var readback = options.FooterSegments!;
            readback[0] = new FooterSegment(FooterSegmentKind.Text, "Readback mutated");
            readback.Add(new FooterSegment(FooterSegmentKind.TotalPages));

            Assert.Equal(2, options.FooterSegments!.Count);
            Assert.Equal("Page ", options.FooterSegments![0].Text);

            var doc = PdfDoc.Create(options)
                .Paragraph(p => p.Text("Footer segment snapshot"));

            string pdfText = Normalize(PdfDocument.Open(new MemoryStream(doc.ToBytes())).GetPage(1).Text);
            Assert.Contains("Page1", pdfText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Mutated", pdfText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Readbackmutated", pdfText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PdfDocCreate_SnapshotsInputOptionsBeforeRendering() {
            var options = new PdfOptions {
                PageWidth = 300,
                PageHeight = 400,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 12
            };

            var doc = PdfDoc.Create(options)
                .Paragraph(p => p.Text("Options snapshot"));

            options.PageWidth = 612;
            options.PageHeight = 792;
            options.DefaultFontSize = 30;
            options.DefaultFont = PdfStandardFont.Courier;

            byte[] bytes = doc.ToBytes();

            using (var pdf = PdfDocument.Open(new MemoryStream(bytes))) {
                var page = pdf.GetPage(1);
                Assert.InRange(page.Width, 299.0, 301.0);
                Assert.InRange(page.Height, 399.0, 401.0);
                double pointSize = page.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                Assert.InRange(pointSize, 11.5, 12.5);
            }
        }

        [Fact]
        public void HeaderText_RendersPageTokensWithSectionLocalPageOptions() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header.AlignRight().Text("Section A {page}/{pages}"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("First page content."))));
                });
                c.Page(page => {
                    page.Header(header => header.AlignLeft().Text("Section B {page}/{pages}"));
                    page.Footer(footer => footer.PageNumberWithTotal());
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Second page content."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("SectionA1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Firstpagecontent", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("SectionB", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("SectionB2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Secondpagecontent", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2/2", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DocumentHeaderFooter_ApplyToTopLevelFlowAndComposePagesCanOverride() {
            var doc = PdfDoc.Create()
                .Margin(PageMargins.Narrow)
                .Header(header => header.AlignLeft().Text("Document header {page}/{pages}"))
                .Footer(footer => footer.AlignCenter().Text(text => text.Text("Document footer ").CurrentPage().Text("/").TotalPages()))
                .Paragraph(p => p.Text("Top flow first."))
                .PageBreak()
                .Paragraph(p => p.Text("Top flow second."));

            doc.Compose(c => c.Page(page => {
                page.Header(header => header.AlignRight().Text("Page header {page}/{pages}"));
                page.Content(content =>
                    content.Column(column =>
                        column.Item().Paragraph(p => p.Text("Composed body."))));
            }));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(3, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("Documentheader1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Documentfooter1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Topflowfirst", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Documentheader2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Documentfooter2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Topflowsecond", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Pageheader3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Documentfooter3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Composedbody", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Documentheader", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderTextBuilder_RendersPageTokensAndOverridesFormat() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Text("Ignored header {page}/{pages}")
                        .Text(text => text.Text("Segment header ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Segment header body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Segmentheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredheader", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderText_RendersLiteralFormatAndOverridesSegmentBuilder() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Text(text => text.Text("Ignored header ").CurrentPage().Text("/").TotalPages())
                        .Text("Literal header {page}/{pages}"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Literal header body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Literalheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredheader", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FooterText_RendersLiteralFormatAndOverridesSegmentBuilder() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Footer(footer => footer
                        .Text(text => text.Text("Ignored footer ").CurrentPage().Text("/").TotalPages())
                        .Text("Literal footer {page}/{pages}"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Literal footer body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Literalfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredfooter", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterTextColor_RendersConfiguredColorsAndResetsFooterAfterBodyText() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Color(new PdfColor(0.1, 0.2, 0.3))
                        .Text("Colored header"));
                    page.Footer(footer => footer
                        .Color(new PdfColor(0.2, 0.3, 0.4))
                        .Text(text => text.Text("Colored footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Color(new PdfColor(0.9, 0.1, 0.1)).Text("Colored body."))));
                });
            });

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Coloredheader", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Coloredbody", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Coloredfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);

            int headerColorIndex = rawPdf.IndexOf("0.1 0.2 0.3 rg", StringComparison.Ordinal);
            int bodyColorIndex = rawPdf.IndexOf("0.9 0.1 0.1 rg", StringComparison.Ordinal);
            int footerColorIndex = rawPdf.IndexOf("0.2 0.3 0.4 rg", StringComparison.Ordinal);

            Assert.True(headerColorIndex >= 0, "The header should use its configured text color.");
            Assert.True(bodyColorIndex > headerColorIndex, "The body should be written after the header.");
            Assert.True(footerColorIndex > bodyColorIndex, "The footer should reset fill color after colored body text.");
        }

        [Fact]
        public void HeaderFooterCompose_RendersConfiguredFontsAndSizes() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Font(PdfStandardFont.HelveticaBold)
                        .FontSize(13)
                        .Text("Typography header"));
                    page.Footer(footer => footer
                        .Font(PdfStandardFont.TimesItalic)
                        .FontSize(8)
                        .Text(text => text.Text("Typography footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Typography body."))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            var letters = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .ToList();

            var headerLetters = letters
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .First()
                .ToList();

            var footerLetters = letters
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderBy(group => group.Key)
                .First()
                .ToList();

            string headerText = Normalize(string.Concat(headerLetters.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)));
            string footerText = Normalize(string.Concat(footerLetters.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)));

            Assert.Equal("Typographyheader", headerText);
            Assert.Equal("Typographyfooter1/1", footerText);
            Assert.Contains(headerLetters, letter => letter.FontName != null && letter.FontName.Contains("Helvetica-Bold", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(footerLetters, letter => letter.FontName != null && letter.FontName.Contains("Times-Italic", StringComparison.OrdinalIgnoreCase));
            Assert.InRange(headerLetters.Select(letter => letter.PointSize).Average(), 12.5, 13.5);
            Assert.InRange(footerLetters.Select(letter => letter.PointSize).Average(), 7.5, 8.5);
        }

        [Fact]
        public void HeaderFooterCompose_RendersConfiguredOffsets() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Size(612, 792);
                    page.Margin(72);
                    page.Header(header => header
                        .Offset(12)
                        .Text("Offset header"));
                    page.Footer(footer => footer
                        .Offset(20)
                        .Text(text => text.Text("Offset footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Offset body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            var groups = pdf.GetPage(1).Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .Select(group => new {
                    Y = group.Key,
                    Text = Normalize(string.Concat(group.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)))
                })
                .ToList();

            var headerLine = Assert.Single(groups, group => group.Text == "Offsetheader");
            var footerLine = Assert.Single(groups, group => group.Text == "Offsetfooter1/1");

            Assert.InRange(headerLine.Y, 731.5, 732.5);
            Assert.InRange(footerLine.Y, 51.5, 52.5);
        }

        [Fact]
        public void HeaderFooterZones_RenderLeftCenterAndRightTextOnOneLine() {
            var doc = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    HeaderFont = PdfStandardFont.Helvetica,
                    FooterFont = PdfStandardFont.Helvetica
                })
                .Size(612, 792)
                .Margin(72)
                .Header(header => header.Zones("HeaderLeft", "HeaderCenter {page}/{pages}", "HeaderRight"))
                .Footer(footer => footer.Zones("FooterLeft", "FooterCenter {page}/{pages}", "FooterRight"))
                .Paragraph(p => p.Text("Zone body."));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            var page = pdf.GetPage(1);
            string pageText = Normalize(page.Text);

            Assert.Contains("HeaderLeft", pageText, StringComparison.Ordinal);
            Assert.Contains("HeaderCenter1/1", pageText, StringComparison.Ordinal);
            Assert.Contains("HeaderRight", pageText, StringComparison.Ordinal);
            Assert.Contains("FooterLeft", pageText, StringComparison.Ordinal);
            Assert.Contains("FooterCenter1/1", pageText, StringComparison.Ordinal);
            Assert.Contains("FooterRight", pageText, StringComparison.Ordinal);

            double headerLeftX = FindWordStartX(page, "HeaderLeft");
            double headerCenterX = FindWordStartX(page, "HeaderCenter");
            double headerRightX = FindWordStartX(page, "HeaderRight");
            double footerLeftX = FindWordStartX(page, "FooterLeft");
            double footerCenterX = FindWordStartX(page, "FooterCenter");
            double footerRightX = FindWordStartX(page, "FooterRight");

            Assert.InRange(headerLeftX, 71.5, 72.5);
            Assert.True(headerCenterX > headerLeftX + 150, $"Expected centered header zone after left zone. Center x: {headerCenterX:0.##}, left x: {headerLeftX:0.##}.");
            Assert.True(headerRightX > headerCenterX + 150, $"Expected right header zone after center zone. Right x: {headerRightX:0.##}, center x: {headerCenterX:0.##}.");
            Assert.InRange(footerLeftX, 71.5, 72.5);
            Assert.True(footerCenterX > footerLeftX + 150, $"Expected centered footer zone after left zone. Center x: {footerCenterX:0.##}, left x: {footerLeftX:0.##}.");
            Assert.True(footerRightX > footerCenterX + 150, $"Expected right footer zone after center zone. Right x: {footerRightX:0.##}, center x: {footerCenterX:0.##}.");
        }

        [Fact]
        public void HeaderFooterZones_AreOverriddenByLaterSingleTextCalls() {
            var doc = PdfDoc.Create()
                .Header(header => header
                    .Zones("Ignored left", "Ignored center", "Ignored right")
                    .Text("Final header {page}/{pages}"))
                .Footer(footer => footer
                    .Zones("Ignored footer left", "Ignored footer center", "Ignored footer right")
                    .Text(text => text.Text("Final footer ").CurrentPage().Text("/").TotalPages()))
                .Paragraph(p => p.Text("Zone override body."));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Finalheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Finalfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignored", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterZones_RejectTextThatWouldOverlap() {
            var headerException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create(new PdfOptions {
                        HeaderFont = PdfStandardFont.Helvetica,
                        HeaderFontSize = 12
                    })
                    .Size(260, 260)
                    .Margin(72)
                    .Header(header => header.Zones(
                        "Very long left header zone",
                        "Very long center header zone",
                        "Very long right header zone"))
                    .Paragraph(p => p.Text("Header zone overlap body."))
                    .ToBytes());

            Assert.Contains("PDF header zone text", headerException.Message, StringComparison.Ordinal);

            var footerException = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create(new PdfOptions {
                        FooterFont = PdfStandardFont.Helvetica,
                        FooterFontSize = 12
                    })
                    .Size(260, 260)
                    .Margin(72)
                    .Footer(footer => footer.Zones(
                        "Very long left footer zone",
                        "Very long center footer zone",
                        "Very long right footer zone"))
                    .Paragraph(p => p.Text("Footer zone overlap body."))
                    .ToBytes());

            Assert.Contains("PDF footer zone text", footerException.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void HeaderFooterZones_CanConfigureFirstAndEvenPageVariants() {
            var doc = PdfDoc.Create(new PdfOptions {
                    HeaderFont = PdfStandardFont.Helvetica,
                    FooterFont = PdfStandardFont.Helvetica
                })
                .Header(header => header
                    .Zones("OddLeft {page}/{pages}", "OddCenter", "OddRight")
                    .FirstPageZones("FirstLeft {page}/{pages}", "FirstCenter", "FirstRight")
                    .EvenPagesZones("EvenLeft {page}/{pages}", "EvenCenter", "EvenRight"))
                .Footer(footer => footer
                    .Zones("OddFooterLeft {page}/{pages}", "OddFooterCenter", "OddFooterRight")
                    .FirstPageZones("FirstFooterLeft {page}/{pages}", "FirstFooterCenter", "FirstFooterRight")
                    .EvenPagesZones("EvenFooterLeft {page}/{pages}", "EvenFooterCenter", "EvenFooterRight"))
                .Paragraph(p => p.Text("First zone variant body."))
                .PageBreak()
                .Paragraph(p => p.Text("Even zone variant body."))
                .PageBreak()
                .Paragraph(p => p.Text("Odd zone variant body."));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(3, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("FirstLeft1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FirstCenter", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FirstRight", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FirstFooterLeft1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FirstFooterCenter", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("FirstFooterRight", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("OddLeft", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("EvenLeft", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("EvenLeft2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("EvenCenter", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("EvenRight", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("EvenFooterLeft2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("EvenFooterCenter", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("EvenFooterRight", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("OddLeft", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("FirstLeft", page2Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("OddLeft3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("OddCenter", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("OddRight", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("OddFooterLeft3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("OddFooterCenter", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("OddFooterRight", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("FirstLeft", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("EvenLeft", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DifferentFirstPageHeaderFooter_UsesFirstPageContentThenRunningContent() {
            var options = new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "Running header {page}/{pages}",
                ShowPageNumbers = true,
                FooterFormat = "Running footer {page}/{pages}",
                DifferentFirstPageHeaderFooter = true,
                FirstPageHeaderFormat = "Cover header {page}/{pages}",
                FirstPageFooterSegments = new System.Collections.Generic.List<FooterSegment> {
                    new FooterSegment(FooterSegmentKind.Text, "Cover footer "),
                    new FooterSegment(FooterSegmentKind.PageNumber),
                    new FooterSegment(FooterSegmentKind.Text, "/"),
                    new FooterSegment(FooterSegmentKind.TotalPages)
                }
            };

            byte[] bytes = PdfDoc.Create(options)
                .Paragraph(p => p.Text("Cover body."))
                .PageBreak()
                .Paragraph(p => p.Text("Running body."))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("Coverheader1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Coverfooter1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Runningheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Runningfooter", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Runningheader2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Runningfooter2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Coverheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Coverfooter", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DifferentFirstPageHeaderFooter_BlankFirstPageSuppressesRunningContent() {
            var options = new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "Running header",
                ShowPageNumbers = true,
                FooterFormat = "Running footer",
                DifferentFirstPageHeaderFooter = true
            };

            byte[] bytes = PdfDoc.Create(options)
                .Paragraph(p => p.Text("Cover body."))
                .PageBreak()
                .Paragraph(p => p.Text("Running body."))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.DoesNotContain("Runningheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Runningfooter", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Runningheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Runningfooter", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DifferentOddAndEvenPagesHeaderFooter_UsesEvenContentAndKeepsFirstPagePrecedence() {
            var options = new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "Odd header {page}/{pages}",
                ShowPageNumbers = true,
                FooterFormat = "Odd footer {page}/{pages}",
                DifferentFirstPageHeaderFooter = true,
                FirstPageHeaderFormat = "First header {page}/{pages}",
                FirstPageFooterFormat = "First footer {page}/{pages}",
                DifferentOddAndEvenPagesHeaderFooter = true,
                EvenPageHeaderFormat = "Even header {page}/{pages}",
                EvenPageFooterSegments = new System.Collections.Generic.List<FooterSegment> {
                    new FooterSegment(FooterSegmentKind.Text, "Even footer "),
                    new FooterSegment(FooterSegmentKind.PageNumber),
                    new FooterSegment(FooterSegmentKind.Text, "/"),
                    new FooterSegment(FooterSegmentKind.TotalPages)
                }
            };

            byte[] bytes = PdfDoc.Create(options)
                .Paragraph(p => p.Text("First body."))
                .PageBreak()
                .Paragraph(p => p.Text("Even body."))
                .PageBreak()
                .Paragraph(p => p.Text("Odd body."))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            Assert.Equal(3, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("Firstheader1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Firstfooter1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evenheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddheader", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Evenheader2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Evenfooter2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddfooter", page2Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Oddheader3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddfooter3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evenheader", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evenfooter", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DifferentOddAndEvenPagesHeaderFooter_BlankEvenPagesSuppressRunningContent() {
            var options = new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "Odd header",
                ShowPageNumbers = true,
                FooterFormat = "Odd footer",
                DifferentOddAndEvenPagesHeaderFooter = true
            };

            byte[] bytes = PdfDoc.Create(options)
                .Paragraph(p => p.Text("Odd body."))
                .PageBreak()
                .Paragraph(p => p.Text("Even body."))
                .PageBreak()
                .Paragraph(p => p.Text("Odd again body."))
                .ToBytes();

            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("Oddheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddfooter", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddfooter", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddheader", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddfooter", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ComposeHeaderFooter_CanConfigureDifferentFirstPageContent() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Text(text => text.Text("Running compose header ").CurrentPage().Text("/").TotalPages())
                        .FirstPageText(text => text.Text("Compose cover header ").CurrentPage().Text("/").TotalPages()));
                    page.Footer(footer => footer
                        .PageNumberWithTotal()
                        .FirstPageText(text => text.Text("Compose cover footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("Compose cover body."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Text("Compose running body."));
                        }));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("Composecoverheader1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Composecoverfooter1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Runningcomposeheader", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Runningcomposeheader2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Composecoverheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Composecoverfooter", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ComposeHeaderFooter_CanConfigureDifferentEvenPageContent() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header
                        .Text(text => text.Text("Odd compose header ").CurrentPage().Text("/").TotalPages())
                        .EvenPagesText(text => text.Text("Even compose header ").CurrentPage().Text("/").TotalPages()));
                    page.Footer(footer => footer
                        .Text(text => text.Text("Odd compose footer ").CurrentPage().Text("/").TotalPages())
                        .EvenPagesText(text => text.Text("Even compose footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("Odd compose body."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Text("Even compose body."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Text("Odd compose body again."));
                        }));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(3, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("Oddcomposeheader1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddcomposefooter1/3", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Evencomposeheader2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Evencomposefooter2/3", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddcomposeheader3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddcomposefooter3/3", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evencomposeheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddcomposeheader", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evencomposeheader", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SectionHeaderFooter_VariantsRestartPerSectionAndPageTokensContinueByDefault() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Section(page => {
                    page.Header(header => header
                        .Text("A odd {page}/{pages}")
                        .FirstPageText("A first {page}/{pages}")
                        .EvenPagesText("A even {page}/{pages}"));
                    page.Footer(footer => footer
                        .Text(text => text.Text("A footer odd ").CurrentPage().Text("/").TotalPages())
                        .FirstPageText(text => text.Text("A footer first ").CurrentPage().Text("/").TotalPages())
                        .EvenPagesText(text => text.Text("A footer even ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("A first body."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Text("A even body."));
                        }));
                });
                c.Section(page => {
                    page.Header(header => header
                        .Text("B odd {page}/{pages}")
                        .FirstPageText("B first {page}/{pages}")
                        .EvenPagesText("B even {page}/{pages}"));
                    page.Footer(footer => footer
                        .Text(text => text.Text("B footer odd ").CurrentPage().Text("/").TotalPages())
                        .FirstPageText(text => text.Text("B footer first ").CurrentPage().Text("/").TotalPages())
                        .EvenPagesText(text => text.Text("B footer even ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("B first body."));
                            column.Item().PageBreak();
                            column.Item().Paragraph(p => p.Text("B even body."));
                        }));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(4, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);
            string page4Text = Normalize(pdf.GetPage(4).Text);

            Assert.Contains("Afirst1/4", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Afooterfirst1/4", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Aeven2/4", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Afootereven2/4", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Bfirst3/4", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Bfooterfirst3/4", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Beven4/4", page4Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Bfootereven4/4", page4Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Bfirst", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Afirst", page3Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SectionHeaderFooter_PageNumberStartChangesVisibleTokensButNotVariants() {
            var doc = PdfDoc.Create();

            doc.Section(section => {
                section.PageNumberStart(5);
                section.Header(header => header
                    .Text("Running {page}/{pages}")
                    .FirstPageText("First {page}/{pages}")
                    .EvenPagesText("Even {page}/{pages}"));
                section.Footer(footer => footer
                    .Text(text => text.Text("Footer running ").CurrentPage().Text("/").TotalPages())
                    .FirstPageText(text => text.Text("Footer first ").CurrentPage().Text("/").TotalPages())
                    .EvenPagesText(text => text.Text("Footer even ").CurrentPage().Text("/").TotalPages()));
                section.Content(content =>
                    content.Column(column => {
                        column.Item().Paragraph(p => p.Text("Started section first body."));
                        column.Item().PageBreak();
                        column.Item().Paragraph(p => p.Text("Started section second body."));
                    }));
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("First5/6", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Footerfirst5/6", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Running5/6", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Even6/6", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Footereven6/6", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("First6/6", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DocumentPageNumberStart_AppliesToComposedPagesAndContinuesAcrossFlows() {
            var doc = PdfDoc.Create()
                .PageNumberStart(5)
                .Header(header => header.Text("Doc {page}/{pages}"));

            doc.Compose(c => {
                c.Page(page => {
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("First composed body."))));
                });
                c.Page(page => {
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Second composed body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("Doc5/6", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Doc6/6", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterPageTokens_UseConfiguredPageNumberStyleForFormatsAndSegments() {
            var doc = PdfDoc.Create()
                .PageNumberStyle(PdfPageNumberStyle.UpperRoman)
                .Header(header => header.Text("Roman header {page}/{pages}"))
                .Footer(footer => footer.Text(text => text.Text("Roman footer ").CurrentPage().Text("/").TotalPages()))
                .Paragraph(p => p.Text("Roman first body."))
                .PageBreak()
                .Paragraph(p => p.Text("Roman second body."));

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("RomanheaderI/II", page1Text, StringComparison.Ordinal);
            Assert.Contains("RomanfooterI/II", page1Text, StringComparison.Ordinal);
            Assert.Contains("RomanheaderII/II", page2Text, StringComparison.Ordinal);
            Assert.Contains("RomanfooterII/II", page2Text, StringComparison.Ordinal);
        }

        [Fact]
        public void SectionHeaderFooter_PageNumberStyleAppliesAfterExplicitStart() {
            var doc = PdfDoc.Create();

            doc.Section(section => {
                section.PageNumberStart(27);
                section.PageNumberStyle(PdfPageNumberStyle.LowerLetter);
                section.Header(header => header
                    .Text("Letter running {page}/{pages}")
                    .FirstPageText("Letter first {page}/{pages}")
                    .EvenPagesText("Letter even {page}/{pages}"));
                section.Footer(footer => footer
                    .Text(text => text.Text("Letter footer running ").CurrentPage().Text("/").TotalPages())
                    .FirstPageText(text => text.Text("Letter footer first ").CurrentPage().Text("/").TotalPages())
                    .EvenPagesText(text => text.Text("Letter footer even ").CurrentPage().Text("/").TotalPages()));
                section.Content(content =>
                    content.Column(column => {
                        column.Item().Paragraph(p => p.Text("Letter section first body."));
                        column.Item().PageBreak();
                        column.Item().Paragraph(p => p.Text("Letter section second body."));
                    }));
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("Letterfirstaa/ab", page1Text, StringComparison.Ordinal);
            Assert.Contains("Letterfooterfirstaa/ab", page1Text, StringComparison.Ordinal);
            Assert.DoesNotContain("Letterrunningaa/ab", page1Text, StringComparison.Ordinal);
            Assert.Contains("Letterevenab/ab", page2Text, StringComparison.Ordinal);
            Assert.Contains("Letterfooterevenab/ab", page2Text, StringComparison.Ordinal);
            Assert.DoesNotContain("Letterfirstab/ab", page2Text, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePagesHaveIndependentOptions() {
            var doc = PdfDoc.Create();
            doc.Compose(c => {
                c.Page(page => {
                    page.Size(PageSizes.A4);
                    page.Margin(36);
                    page.DefaultTextStyle(style => style.Font(PdfStandardFont.Helvetica).FontSize(14));
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("First page body."))));
                });
                c.Page(page => {
                    page.Size(PageSizes.Letter);
                    page.Margin(18, 24, 30, 36);
                    page.DefaultTextStyle(style => style.Font(PdfStandardFont.TimesRoman).FontSize(11));
                    page.Footer(f => f.PageNumber());
                    page.Content(content =>
                        content.Column(col =>
                            col.Item().Paragraph(p => p.Text("Second page body."))));
                });
            });

            var bytes = doc.ToBytes();
            Assert.NotEmpty(bytes);

            using (var pdf = PdfDocument.Open(new MemoryStream(bytes))) {
                Assert.Equal(2, pdf.NumberOfPages);
                var page1 = pdf.GetPage(1);
                var page2 = pdf.GetPage(2);
                Assert.InRange(page1.Width, 594.0, 596.0); // A4 width
                Assert.InRange(page1.Height, 841.0, 843.0); // A4 height
                Assert.InRange(page2.Width, 611.0, 613.0); // Letter width
                Assert.InRange(page2.Height, 791.0, 793.0); // Letter height

                var page1Text = string.Concat(page1.Letters.Select(l => l.Value));
                var page2Text = string.Concat(page2.Letters.Select(l => l.Value));
                var page1TextNormalized = Normalize(page1Text);
                var page2TextNormalized = Normalize(page2Text);
                Assert.Contains("Firstpagebody", page1TextNormalized, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("Secondpagebody", page1TextNormalized, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Secondpagebody", page2TextNormalized, StringComparison.OrdinalIgnoreCase);

                double page1Size = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                double page2Size = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                Assert.InRange(page1Size, 13.5, 14.5);
                Assert.InRange(page2Size, 10.5, 11.5);
            }
        }

        [Fact]
        public void PdfDocPage_CreatesPageScopedFlowWithoutComposeWrapper() {
            var doc = PdfDoc.Create()
                .Size(PageSizes.Letter)
                .Margin(PageMargins.Normal)
                .Header(header => header.Text("Document header {page}/{pages}"))
                .Page(page => {
                    page.Size(PageSizes.A5);
                    page.Margin(PageMargins.Narrow);
                    page.Header(header => header.Text("Small page {page}/{pages}"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Scoped page body."))));
                })
                .Page(page => {
                    page.Landscape();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Inherited landscape body."))));
                });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            string page1Text = Normalize(page1.Text);
            string page2Text = Normalize(page2.Text);

            Assert.InRange(page1.Width, 419.0, 421.0);
            Assert.InRange(page1.Height, 594.0, 596.0);
            Assert.InRange(FindWordStartX(page1, "Scoped"), 35.5, 36.5);
            Assert.Contains("Smallpage1/2", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Scopedpagebody", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Documentheader", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.InRange(page2.Width, 791.0, 793.0);
            Assert.InRange(page2.Height, 611.0, 613.0);
            Assert.InRange(FindWordStartX(page2, "Inherited"), 71.5, 72.5);
            Assert.Contains("Documentheader2/2", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Inheritedlandscapebody", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PdfDocSection_CreatesSectionScopedFlowAcrossPhysicalPages() {
            var doc = PdfDoc.Create()
                .Section(section => {
                    section.Size(220, 170);
                    section.Margin(24);
                    section.Header(header => header.Text("Small section {page}/{pages}"));
                    section.Content(content =>
                        content.Column(column => {
                            for (int i = 1; i <= 18; i++) {
                                int item = i;
                                column.Item().Paragraph(p => p.Text("Small section item " + item.ToString("D2", System.Globalization.CultureInfo.InvariantCulture)));
                            }
                        }));
                })
                .Section(section => {
                    section.Size(PageSizes.A5.Landscape());
                    section.Margin(PageMargins.Narrow);
                    section.Header(header => header.Text("Wide section {page}/{pages}"));
                    section.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Wide section body."))));
                });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.True(pdf.NumberOfPages >= 3, "The first section should flow across multiple physical pages.");

            int widePageNumber = Enumerable.Range(1, pdf.NumberOfPages)
                .First(pageNumber => Normalize(pdf.GetPage(pageNumber).Text).Contains("Widesectionbody", StringComparison.OrdinalIgnoreCase));

            Assert.True(widePageNumber > 1);
            for (int pageNumber = 1; pageNumber < widePageNumber; pageNumber++) {
                var smallPage = pdf.GetPage(pageNumber);
                string text = Normalize(smallPage.Text);
                Assert.InRange(smallPage.Width, 219.0, 221.0);
                Assert.InRange(smallPage.Height, 169.0, 171.0);
                Assert.Contains("Smallsection" + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + pdf.NumberOfPages.ToString(System.Globalization.CultureInfo.InvariantCulture), text, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("Widesection", text, StringComparison.OrdinalIgnoreCase);
            }

            var widePage = pdf.GetPage(widePageNumber);
            string wideText = Normalize(widePage.Text);
            Assert.InRange(widePage.Width, 594.0, 596.0);
            Assert.InRange(widePage.Height, 419.0, 421.0);
            Assert.Contains("Widesection" + widePageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + pdf.NumberOfPages.ToString(System.Globalization.CultureInfo.InvariantCulture), wideText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Widesectionbody", wideText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PageSizeOrientationHelpers_ReturnExpectedDimensionsAndRejectInvalidOrientation() {
            PageSize portrait = PageSizes.A4.Portrait();
            PageSize landscape = PageSizes.A4.Landscape();

            Assert.InRange(portrait.Width, 594.0, 596.0);
            Assert.InRange(portrait.Height, 841.0, 843.0);
            Assert.InRange(landscape.Width, 841.0, 843.0);
            Assert.InRange(landscape.Height, 594.0, 596.0);
            Assert.Equal(landscape.Width, new PageSize(842, 595).Landscape().Width);
            Assert.Equal(landscape.Height, new PageSize(842, 595).Landscape().Height);

            var exception = Assert.Throws<ArgumentException>(() =>
                PageSizes.A4.WithOrientation((PdfPageOrientation)99));

            Assert.Equal("orientation", exception.ParamName);
            Assert.Contains("PDF page orientation must be Portrait or Landscape.", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePageOrientation_PreservesPageSizeAndRendersIndependentPageGeometry() {
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Size(PageSizes.A4);
                    page.Landscape();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Landscape page body."))));
                });
                c.Page(page => {
                    page.Size(PageSizes.A4.Landscape());
                    page.Portrait();
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Portrait page body."))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            var landscapePage = pdf.GetPage(1);
            var portraitPage = pdf.GetPage(2);

            Assert.InRange(landscapePage.Width, 841.0, 843.0);
            Assert.InRange(landscapePage.Height, 594.0, 596.0);
            Assert.InRange(portraitPage.Width, 594.0, 596.0);
            Assert.InRange(portraitPage.Height, 841.0, 843.0);
            Assert.Contains("Landscapepagebody", Normalize(landscapePage.Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Portraitpagebody", Normalize(portraitPage.Text), StringComparison.OrdinalIgnoreCase);

            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.Orientation((PdfPageOrientation)99))));

            Assert.Equal("orientation", exception.ParamName);
        }

        [Fact]
        public void ComposePage_DefaultParagraphStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfParagraphStyle {
                FirstLineIndent = 24,
                SpacingAfter = 0
            };
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.DefaultParagraphStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageOneFirst").LineBreak().Text("PageOneSecond"))));
                });

                style.FirstLineIndent = 0;

                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageTwoFirst").LineBreak().Text("PageTwoSecond"))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneFirstX = FindWordStartX(page1, "PageOneFirst");
            double pageOneSecondX = FindWordStartX(page1, "PageOneSecond");
            double pageTwoFirstX = FindWordStartX(page2, "PageTwoFirst");
            double pageTwoSecondX = FindWordStartX(page2, "PageTwoSecond");

            Assert.True(pageOneFirstX - pageOneSecondX >= 22, $"Expected page default paragraph style to indent first page only. First x: {pageOneFirstX:0.##}, second x: {pageOneSecondX:0.##}.");
            Assert.InRange(System.Math.Abs(pageTwoFirstX - pageTwoSecondX), 0, 2);
        }

        [Fact]
        public void ComposePage_DefaultTextStyleObjectAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfTextStyle {
                Font = PdfStandardFont.Helvetica,
                FontSize = 16,
                Color = PdfColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDoc.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageOneTextStyle"))));
                });

                style.FontSize = 8;
                style.Color = PdfColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageTwoTextStyle"))));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();

            Assert.InRange(pageOneSize, 15.5, 16.5);
            Assert.InRange(pageTwoSize, 9.5, 10.5);
        }

        [Fact]
        public void ComposePage_DefaultHeadingStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfHeadingStyle {
                FontSize = 13,
                Color = PdfColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultHeadingStyle(2, style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().H2("PageOneHeading")));
                });

                style.FontSize = 30;
                style.Color = PdfColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().H2("PageTwoHeading")));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.InRange(pageOneSize, 12.5, 13.5);
            Assert.InRange(pageTwoSize, 17.5, 18.5);
            Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultListStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfListStyle {
                FontSize = 13,
                LeftIndent = 14,
                Color = PdfColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultListStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Bullets(new[] { "PageOneList" })));
                });

                style.FontSize = 30;
                style.LeftIndent = 0;
                style.Color = PdfColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Bullets(new[] { "PageTwoList" })));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0]) && l.Value != "•").Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0]) && l.Value != "•").Select(l => l.PointSize).First();
            double pageOneBulletX = page1.Letters.First(l => l.Value == "•").StartBaseLine.X;
            double pageTwoBulletX = page2.Letters.First(l => l.Value == "•").StartBaseLine.X;
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.InRange(pageOneSize, 12.5, 13.5);
            Assert.InRange(pageTwoSize, 9.5, 10.5);
            Assert.InRange(pageOneBulletX, 43, 45);
            Assert.InRange(pageTwoBulletX, 29.5, 30.5);
            Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultPanelStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PanelStyle {
                PaddingX = 16,
                MaxWidth = 180,
                Align = PdfAlign.Center,
                Background = PdfColor.FromRgb(240, 248, 255)
            };
            var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 240,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultPanelStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().PanelParagraph(p => p.Text("PageOnePanel"))));
                });

                style.PaddingX = 2;
                style.MaxWidth = 300;
                style.Align = PdfAlign.Right;
                style.Background = PdfColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().PanelParagraph(p => p.Text("PageTwoPanel"))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));

            double pageOneX = FindWordStartX(pdf.GetPage(1), "PageOnePanel");
            double pageTwoX = FindWordStartX(pdf.GetPage(2), "PageTwoPanel");
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.InRange(pageOneX, 105, 107);
            Assert.InRange(pageTwoX, 35, 37);
            Assert.Contains("0.941 0.973 1 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultHorizontalRuleStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfHorizontalRuleStyle {
                Thickness = 2,
                Color = PdfColor.FromRgb(10, 20, 30),
                SpacingBefore = 3,
                SpacingAfter = 16
            };
            var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(20);
                    page.DefaultHorizontalRuleStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .HR()
                                .Paragraph(p => p.Text("PageOneRule"))));
                });

                style.Thickness = 6;
                style.Color = PdfColor.FromRgb(200, 10, 10);
                style.SpacingAfter = 0;

                c.Page(page => {
                    page.Margin(20);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .HR()
                                .Paragraph(p => p.Text("PageTwoRule"))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));

            double pageOneY = FindWordStartY(pdf.GetPage(1), "PageOneRule");
            double pageTwoY = FindWordStartY(pdf.GetPage(2), "PageTwoRule");
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.True(pageTwoY - pageOneY >= 7, $"Expected the page-scoped rule style to push only page-one content down. Page one y: {pageOneY:0.##}, page two y: {pageTwoY:0.##}.");
            Assert.Contains("0.039 0.078 0.118 RG", rawPdf, StringComparison.Ordinal);
            Assert.DoesNotContain("0.784 0.039 0.039 RG", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultImageStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            byte[] png = CreateMinimalRgbPng();
            var style = new PdfImageStyle {
                Align = PdfAlign.Center,
                SpacingBefore = 4,
                SpacingAfter = 12
            };
            var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(20);
                    page.DefaultImageStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .Image(png, 24, 24)
                                .Paragraph(p => p.Text("PageOneImage"))));
                });

                style.Align = PdfAlign.Right;
                style.SpacingAfter = 0;

                c.Page(page => {
                    page.Margin(20);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .Image(png, 24, 24)
                                .Paragraph(p => p.Text("PageTwoImage"))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            string rawPdf = Encoding.ASCII.GetString(bytes);
            double pageOneY = FindWordStartY(pdf.GetPage(1), "PageOneImage");
            double pageTwoY = FindWordStartY(pdf.GetPage(2), "PageTwoImage");

            Assert.Contains("q\n24 0 0 24 108 136 cm\n/Im1 Do\nQ", rawPdf);
            Assert.Contains("q\n24 0 0 24 20 136 cm\n/Im", rawPdf);
            Assert.True(pageTwoY - pageOneY >= 10, $"Expected the page-scoped image spacing to push only page-one content down. Page one y: {pageOneY:0.##}, page two y: {pageTwoY:0.##}.");
        }

        [Fact]
        public void ComposePage_DefaultDrawingStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfDrawingStyle {
                Align = PdfAlign.Center,
                SpacingBefore = 4,
                SpacingAfter = 12
            };
            var shape = OfficeShape.Rectangle(40, 20);
            shape.FillColor = OfficeColor.WhiteSmoke;
            var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(20);
                    page.DefaultDrawingStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .Shape(shape)
                                .Paragraph(p => p.Text("PageOneDrawing"))));
                });

                style.Align = PdfAlign.Right;
                style.SpacingAfter = 0;

                c.Page(page => {
                    page.Margin(20);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item()
                                .Shape(shape)
                                .Paragraph(p => p.Text("PageTwoDrawing"))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfDocument.Open(new MemoryStream(bytes));
            string rawPdf = Encoding.ASCII.GetString(bytes);
            double pageOneY = FindWordStartY(pdf.GetPage(1), "PageOneDrawing");
            double pageTwoY = FindWordStartY(pdf.GetPage(2), "PageTwoDrawing");

            Assert.Contains("100 140 40 20 re f", rawPdf, StringComparison.Ordinal);
            Assert.Contains("20 140 40 20 re f", rawPdf, StringComparison.Ordinal);
            Assert.True(pageTwoY - pageOneY >= 10, $"Expected the page-scoped drawing spacing to push only page-one content down. Page one y: {pageOneY:0.##}, page two y: {pageTwoY:0.##}.");
        }

        [Fact]
        public void ComposePage_ThemeAppliesOnlyToThatPageAndSnapshotsInput() {
            var textStyle = new PdfTextStyle {
                Font = PdfStandardFont.Helvetica,
                FontSize = 16,
                Color = PdfColor.FromRgb(10, 20, 30)
            };
            var tableStyle = TableStyles.Minimal();
            tableStyle.CellPaddingX = 22;
            var theme = new PdfTheme {
                TextStyle = textStyle,
                TableStyle = tableStyle
            };
            var doc = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.Theme(theme);
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("ThemePageOne"));
                            column.Item().Table(new[] {
                                new[] { "ThemeOneTable", "Value" },
                                new[] { "Row", "1" }
                            });
                        }));
                });

                textStyle.FontSize = 8;
                tableStyle.CellPaddingX = 0;
                theme.TextStyle = new PdfTextStyle {
                    Font = PdfStandardFont.Helvetica,
                    FontSize = 8
                };

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("ThemePageTwo"));
                            column.Item().Table(new[] {
                                new[] { "ThemeTwoTable", "Value" },
                                new[] { "Row", "2" }
                            });
                        }));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            double pageOneTableX = FindWordStartX(page1, "ThemeOneTable");
            double pageTwoTableX = FindWordStartX(page2, "ThemeTwoTable");

            Assert.InRange(pageOneSize, 15.5, 16.5);
            Assert.InRange(pageTwoSize, 9.5, 10.5);
            Assert.True(pageOneTableX - 30 >= 20, $"Expected page theme table style padding to apply to page one. Marker x: {pageOneTableX:0.##}.");
            Assert.InRange(pageTwoTableX - 30, 4, 8);
        }

        [Fact]
        public void ComposePage_DefaultTableStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = TableStyles.Minimal();
            style.CellPaddingX = 22;
            var doc = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTableStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Table(new[] {
                                new[] { "PageOnePad", "Value" },
                                new[] { "Row", "1" }
                            })));
                });

                style.CellPaddingX = 0;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Table(new[] {
                                new[] { "PageTwoPad", "Value" },
                                new[] { "Row", "2" }
                            })));
                });
            });

            using var pdf = PdfDocument.Open(new MemoryStream(doc.ToBytes()));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneX = FindWordStartX(page1, "PageOnePad");
            double pageTwoX = FindWordStartX(page2, "PageTwoPad");

            Assert.True(pageOneX - 30 >= 20, $"Expected page default table style padding to apply to page one. Marker x: {pageOneX:0.##}.");
            Assert.InRange(pageTwoX - 30, 4, 8);
        }

        [Fact]
        public void ComposePage_DefaultTableStyleRejectsUnsupportedWordStyleName() {
            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create().Compose(c => c.Page(page => page.DefaultTableStyle("Missing Table Style"))));

            Assert.Equal("styleName", exception.ParamName);
            Assert.Contains("Unsupported Word table style", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_ExposesReadOnlyPageBlockCollection() {
            var doc = PdfDoc.Create();

            doc.Compose(c =>
                c.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(paragraph => paragraph.Text("Owned page content."))))));

            var pageBlock = Assert.IsType<PageBlock>(Assert.Single(doc.Blocks));

            Assert.False(pageBlock.Blocks is System.Collections.Generic.List<IPdfBlock>);
            Assert.Single(pageBlock.Blocks);
            Assert.IsType<RichParagraphBlock>(pageBlock.Blocks[0]);
        }

        private static string Normalize(string text) {
            return new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
        }

        private static void AssertMargins(PageMargins margins, double left, double top, double right, double bottom) {
            Assert.Equal(left, margins.Left, 6);
            Assert.Equal(top, margins.Top, 6);
            Assert.Equal(right, margins.Right, 6);
            Assert.Equal(bottom, margins.Bottom, 6);
        }

        private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
            var lines = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => System.Math.Round(letter.StartBaseLine.Y, 1));

            foreach (var line in lines) {
                var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
                string text = string.Concat(ordered.Select(letter => letter.Value));
                int index = text.IndexOf(word, StringComparison.Ordinal);
                if (index >= 0) {
                    return ordered[index].StartBaseLine.X;
                }
            }

            throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
        }

        private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
            var lines = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => System.Math.Round(letter.StartBaseLine.Y, 1));

            foreach (var line in lines) {
                var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
                string text = string.Concat(ordered.Select(letter => letter.Value));
                int index = text.IndexOf(word, StringComparison.Ordinal);
                if (index >= 0) {
                    return ordered[index].StartBaseLine.Y;
                }
            }

            throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
        }

        private static byte[] CreateMinimalRgbPng() {
            return new byte[] {
                137, 80, 78, 71, 13, 10, 26, 10,
                0, 0, 0, 13,
                73, 72, 68, 82,
                0, 0, 0, 1,
                0, 0, 0, 1,
                8, 2, 0, 0, 0,
                0, 0, 0, 0,
                0, 0, 0, 12,
                73, 68, 65, 84,
                0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
                0, 0, 0, 0,
                0, 0, 0, 0,
                73, 69, 78, 68,
                0, 0, 0, 0
            };
        }
    }
}
