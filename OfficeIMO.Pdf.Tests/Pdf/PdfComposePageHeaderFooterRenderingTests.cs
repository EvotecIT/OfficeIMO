using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfDocumentBuilderPageOptionsTests {
        [Fact]
        public void PdfDocumentCreate_SnapshotsInputOptionsBeforeRendering() {
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

            var doc = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Options snapshot"));

            options.PageWidth = 612;
            options.PageHeight = 792;
            options.DefaultFontSize = 30;
            options.DefaultFont = PdfStandardFont.Courier;

            byte[] bytes = doc.ToBytes();

            using (var pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
                var page = pdf.GetPage(1);
                Assert.InRange(page.Width, 299.0, 301.0);
                Assert.InRange(page.Height, 399.0, 401.0);
                double pointSize = page.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
                Assert.InRange(pointSize, 11.5, 12.5);
            }
        }

        [Fact]
        public void HeaderText_RendersPageTokensWithSectionLocalPageOptions() {
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create()
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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Segmentheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredheader", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderText_RendersLiteralFormatAndOverridesSegmentBuilder() {
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Literalheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredheader", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void FooterText_RendersLiteralFormatAndOverridesSegmentBuilder() {
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Literalfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignoredfooter", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterTextColor_RendersConfiguredColorsAndResetsFooterAfterBodyText() {
            var doc = PdfDocument.Create();

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
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
        public void HeaderFooterTextColor_DefaultFooterResetsAfterBodyText() {
            var doc = PdfDocument.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Footer(footer => footer
                        .Text(text => text.Text("Default footer ").CurrentPage().Text("/").TotalPages()));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Color(new PdfColor(0.1, 0.7, 0.2)).Text("Green body."))));
                });
            });

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Greenbody", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Defaultfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);

            int bodyColorIndex = rawPdf.IndexOf("0.1 0.7 0.2 rg", StringComparison.Ordinal);
            int footerColorIndex = rawPdf.IndexOf("0 0 0 rg", bodyColorIndex + 1, StringComparison.Ordinal);

            Assert.True(bodyColorIndex >= 0, "The body should use its configured text color.");
            Assert.True(footerColorIndex > bodyColorIndex, "The default footer should reset fill color after colored body text.");
        }

        [Fact]
        public void HeaderFooterText_RendersLineBreaksOnSeparateLines() {
            var doc = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 50,
                MarginRight = 50,
                MarginTop = 60,
                MarginBottom = 60,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            });

            doc.Compose(c => {
                c.Page(page => {
                    page.Header(header => header.Zones("HeaderLineOne\nHeaderLineTwo", null, "RightLineOne\nRightLineTwo"));
                    page.Footer(footer => footer.AlignCenter().Text("FooterLineOne\nFooterLineTwo"));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("Multiline header footer body."))));
                });
            });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            var page = pdf.GetPage(1);
            string text = page.Text;

            Assert.Contains("HeaderLineOne", text);
            Assert.Contains("HeaderLineTwo", text);
            Assert.Contains("RightLineOne", text);
            Assert.Contains("RightLineTwo", text);
            Assert.Contains("FooterLineOne", text);
            Assert.Contains("FooterLineTwo", text);

            double firstHeaderY = FindWordStartY(page, "HeaderLineOne");
            double secondHeaderY = FindWordStartY(page, "HeaderLineTwo");
            double firstFooterY = FindWordStartY(page, "FooterLineOne");
            double secondFooterY = FindWordStartY(page, "FooterLineTwo");

            Assert.True(firstHeaderY > secondHeaderY + 5D, $"Expected the second header line below the first. First y: {firstHeaderY:0.##}, second y: {secondHeaderY:0.##}.");
            Assert.True(firstFooterY > secondFooterY + 5D, $"Expected the second footer line below the first. First y: {firstFooterY:0.##}, second y: {secondFooterY:0.##}.");
        }

        [Fact]
        public void HeaderFooterCompose_RendersConfiguredFontsAndSizes() {
            var doc = PdfDocument.Create();

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
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    HeaderFont = PdfStandardFont.Helvetica,
                    FooterFont = PdfStandardFont.Helvetica
                })
                .Size(612, 792)
                .Margin(72)
                .Header(header => header.Zones("HeaderLeft", "HeaderCenter {page}/{pages}", "HeaderRight"))
                .Footer(footer => footer.Zones("FooterLeft", "FooterCenter {page}/{pages}", "FooterRight"))
                .Paragraph(p => p.Text("Zone body."));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create()
                .Header(header => header
                    .Zones("Ignored left", "Ignored center", "Ignored right")
                    .Text("Final header {page}/{pages}"))
                .Footer(footer => footer
                    .Zones("Ignored footer left", "Ignored footer center", "Ignored footer right")
                    .Text(text => text.Text("Final footer ").CurrentPage().Text("/").TotalPages()))
                .Paragraph(p => p.Text("Zone override body."));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            string pageText = Normalize(pdf.GetPage(1).Text);

            Assert.Contains("Finalheader1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Finalfooter1/1", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Ignored", pageText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterZones_PreserveOverflowAndReportLayoutDiagnostics() {
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 12,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 12
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Size(260, 260)
                .Margin(72)
                .Header(header => header.Zones(
                    "LeftHeaderOverflowMarker",
                    "CenterHeaderOverflowMarker",
                    "RightHeaderOverflowMarker"))
                .Footer(footer => footer.Zones(
                    "LeftFooterOverflowMarker",
                    "CenterFooterOverflowMarker",
                    "RightFooterOverflowMarker"))
                .Paragraph(p => p.Text("Header footer overflow body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string pageText = Normalize(pdf.GetPage(1).Text);
            Assert.Contains("LeftHeaderOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("CenterHeaderOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("RightHeaderOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("LeftFooterOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("CenterFooterOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("RightFooterOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Header" &&
                warning.Severity == PdfConversionWarningSeverity.Information &&
                warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.Overflow);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Footer" &&
                warning.Severity == PdfConversionWarningSeverity.Information &&
                warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.Overflow);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterZoneOverlap" &&
                warning.Source == "Header");
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterZoneOverlap" &&
                warning.Source == "Footer");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterSingleText_PreservesOverflowWithInformationalDiagnostics() {
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 12,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 12
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Size(260, 260)
                .Margin(72)
                .Header(header => header.AlignCenter().Text("SingleHeaderOverflowMarker"))
                .Footer(footer => footer.AlignRight().Text(text => text.Text("SingleFooterOverflowMarker")))
                .Paragraph(p => p.Text("Single text overflow body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string pageText = Normalize(pdf.GetPage(1).Text);
            Assert.Contains("SingleHeaderOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("SingleFooterOverflowMarker", pageText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Header" &&
                warning.Severity == PdfConversionWarningSeverity.Information);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Footer" &&
                warning.Severity == PdfConversionWarningSeverity.Information);
            Assert.False(report.HasLoss);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void HeaderFooterTextBeyondPhysicalPageReportsVerticalClipping(bool useZones) {
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 40,
                HeaderOffsetY = 40,
                FooterOffsetY = 40,
                HeaderFontSize = 12,
                FooterFontSize = 12
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            PdfDocument document = PdfDocument.Create(options);
            if (useZones) {
                document
                    .Header(header => header.Zones("HeaderVerticalClip", null, null))
                    .Footer(footer => footer.Zones("FooterVerticalClip", null, null));
            } else {
                document
                    .Header(header => header.Text("HeaderVerticalClip"))
                    .Footer(footer => footer.Text("FooterVerticalClip"));
            }

            byte[] bytes = document
                .Paragraph(paragraph => paragraph.Text("Vertical clipping diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            PdfConversionWarning[] clippingWarnings = report.Warnings
                .Where(warning => warning.Code == "HeaderFooterPageBoundsClipped")
                .ToArray();
            Assert.Contains(clippingWarnings, warning =>
                warning.Source == "Header" &&
                (warning.LayoutDiagnostic?.Y ?? 0D) + (warning.LayoutDiagnostic?.Height ?? 0D) > options.PageHeight &&
                warning.LayoutDiagnostic?.Height > 0D);
            Assert.Contains(clippingWarnings, warning =>
                warning.Source == "Footer" &&
                warning.LayoutDiagnostic?.Y < 0D &&
                warning.LayoutDiagnostic.Height > 0D);
            Assert.True(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterRichTextDecorationsUsePaintedVerticalBoundsForClippingDiagnostics() {
            var backgroundReport = new PdfConversionReport();
            var backgroundOptions = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 10,
                MarginBottom = 40,
                HeaderOffsetY = 1,
                HeaderFontSize = 12
            }.ReportDiagnosticsTo(backgroundReport, "OfficeIMO.Tests");

            byte[] backgroundBytes = PdfDocument.Create(backgroundOptions)
                .Header(header => header.Text(text => text.Run(PdfTextRun.Normal(
                    "HeaderBackgroundClip",
                    fontSize: 12,
                    backgroundColor: PdfColor.FromRgb(255, 255, 0)))))
                .Paragraph(paragraph => paragraph.Text("Header decoration diagnostic body"))
                .ToBytes();

            var underlineReport = new PdfConversionReport();
            var underlineOptions = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 40,
                MarginBottom = 1,
                FooterOffsetY = 0,
                FooterFontSize = 1
            }.ReportDiagnosticsTo(underlineReport, "OfficeIMO.Tests");

            byte[] underlineBytes = PdfDocument.Create(underlineOptions)
                .Footer(footer => footer.Text(text => text.Run(PdfTextRun.Underlined("FooterUnderlineClip", fontSize: 1))))
                .Paragraph(paragraph => paragraph.Text("Footer decoration diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(backgroundBytes);
            Assert.Contains(backgroundReport.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header" &&
                (warning.LayoutDiagnostic?.Y ?? 0D) + (warning.LayoutDiagnostic?.Height ?? 0D) > backgroundOptions.PageHeight);
            Assert.True(backgroundReport.HasLoss);

            Assert.NotEmpty(underlineBytes);
            Assert.Contains(underlineReport.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Footer" &&
                warning.LayoutDiagnostic?.Y < 0D);
            Assert.True(underlineReport.HasLoss);
        }

        [Fact]
        public void HeaderFooterZones_CanConfigureFirstAndEvenPageVariants() {
            var doc = PdfDocument.Create(new PdfOptions {
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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
        public void HeaderFooterImages_RenderInsideMarginAreasWithoutImplicitFooterText() {
            byte[] png = CreateMinimalRgbPng();
            var doc = PdfDocument.Create(new PdfOptions {
                    PageWidth = 300,
                    PageHeight = 200,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 40,
                    MarginBottom = 40,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .Header(header => header.Image(png, 24, 12, PdfAlign.Left).Text("HeaderImageText"))
                .Footer(footer => footer.Image(png, 30, 10, PdfAlign.Center))
                .Paragraph(paragraph => paragraph.Text("Header footer image body"));

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.Contains("24 0 0 12 30 166 cm", rawPdf);
            Assert.Contains("30 0 0 10 135 22 cm", rawPdf);
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string text = pdf.GetPage(1).Text;
            Assert.Contains("HeaderImageText", text);
            Assert.Contains("Header footer image body", text);
            Assert.DoesNotContain("Page 1", text, StringComparison.Ordinal);
        }

        [Fact]
        public void HeaderFooterImages_SharingAnAlignmentFormOneNonOverlappingGroupWithText() {
            byte[] png = CreateMinimalRgbPng();
            var doc = PdfDocument.Create(new PdfOptions {
                    PageWidth = 300,
                    PageHeight = 200,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 40,
                    MarginBottom = 40,
                    HeaderFontSize = 9
                })
                .Header(header => header
                    .Zones(null, "Centered header", null)
                    .Image(png, 24, 12, PdfAlign.Center))
                .Paragraph(paragraph => paragraph.Text("Body"));

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);
            System.Text.RegularExpressions.Match imagePlacement = System.Text.RegularExpressions.Regex.Match(
                rawPdf,
                @"24 0 0 12 (?<x>-?\d+(?:\.\d+)?) 166 cm",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            Assert.True(imagePlacement.Success, "Expected a centered 24-by-12 header image placement.");
            double imageX = double.Parse(imagePlacement.Groups["x"].Value, System.Globalization.CultureInfo.InvariantCulture);

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            var headerLetters = pdf.GetPage(1).Letters
                .Where(letter => letter.StartBaseLine.Y > 170D)
                .ToList();
            Assert.NotEmpty(headerLetters);

            double textLeft = headerLetters.Min(letter => letter.StartBaseLine.X);
            double textRight = headerLetters.Max(letter => letter.EndBaseLine.X);
            Assert.True(imageX >= textRight + 3.9D, "Aligned header text and images must not overlap.");
            Assert.InRange(Math.Abs((textLeft + imageX + 24D) / 2D - 150D), 0D, 0.1D);
        }

        [Fact]
        public void HeaderFooterImagesAndShapes_PreserveMarginOverflowWithDiagnostics() {
            byte[] png = CreateMinimalRgbPng();
            OfficeShape footerShape = OfficeShape.Rectangle(160, 12);
            footerShape.FillColor = OfficeColor.Blue;
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 80,
                MarginRight = 80,
                MarginTop = 50,
                MarginBottom = 50,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Image(png, 150, 14, PdfAlign.Center))
                .Footer(footer => footer.Shape(footerShape, PdfAlign.Right))
                .Paragraph(paragraph => paragraph.Text("Header footer visual overflow body"))
                .ToBytes();

            string rawPdf = Encoding.ASCII.GetString(bytes);
            Assert.Contains("150 0 0 14 55 174 cm", rawPdf, StringComparison.Ordinal);
            Assert.Contains("20 32 160 12 re", rawPdf, StringComparison.Ordinal);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Header" &&
                warning.Severity == PdfConversionWarningSeverity.Information &&
                warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.Overflow);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterContentOverflow" &&
                warning.Source == "Footer" &&
                warning.Severity == PdfConversionWarningSeverity.Information &&
                warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.Overflow);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterTransformedShapeUsesVisibleBoundsForClippingDiagnostics() {
            OfficeShape shape = OfficeShape.Rectangle(24, 12);
            shape.FillColor = OfficeColor.Blue;
            shape.Transform = OfficeTransform.Scale(1.5D, 1D).Then(OfficeTransform.Translate(220D, 0D));
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Shape(shape, PdfAlign.Left))
                .Paragraph(paragraph => paragraph.Text("Transformed header shape diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic?.Kind);
            Assert.True((warning.LayoutDiagnostic?.X ?? 0D) + (warning.LayoutDiagnostic?.Width ?? 0D) > options.PageWidth);
            Assert.InRange(warning.LayoutDiagnostic?.Width ?? 0D, 35.999D, 36.001D);
            Assert.True(report.HasLoss);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void HeaderFooterShapeClipBoundsConstrainBaseGeometryDiagnostics(bool transformed) {
            OfficeShape shape = OfficeShape.Rectangle(300, 12);
            shape.FillColor = OfficeColor.Blue;
            shape.ClipPath = OfficeClipPath.Rectangle(20, 12);
            if (transformed) {
                shape.Transform = OfficeTransform.Translate(5D, 0D);
            }

            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Shape(shape, PdfAlign.Left))
                .Paragraph(paragraph => paragraph.Text("Header shape clip diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterShapeShadowUsesVisibleBoundsForClippingDiagnostics() {
            OfficeShape shape = OfficeShape.Rectangle(24, 12);
            shape.FillColor = OfficeColor.Blue;
            shape.ClipPath = OfficeClipPath.Rectangle(12, 12);
            shape.Transform = OfficeTransform.Identity;
            shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.5D, 45D, 0D);
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                CompressContentStreams = false,
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Shape(shape, PdfAlign.Right))
                .Paragraph(paragraph => paragraph.Text("Header shape shadow diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            string rawPdf = Encoding.ASCII.GetString(bytes);
            Assert.Equal(1, CountOccurrences(rawPdf, "0 0 12 12 re W n"));
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic?.Kind);
            Assert.InRange(warning.LayoutDiagnostic?.X ?? 0D, 195.999D, 196.001D);
            Assert.InRange(warning.LayoutDiagnostic?.Width ?? 0D, 68.999D, 69.001D);
            Assert.True(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterShapeStrokeUsesVisibleBoundsForClippingDiagnostics() {
            OfficeShape shape = OfficeShape.Rectangle(24, 12);
            shape.FillColor = OfficeColor.Blue;
            shape.StrokeColor = OfficeColor.Black;
            shape.StrokeWidth = 4;
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 1,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Shape(shape, PdfAlign.Left))
                .Paragraph(paragraph => paragraph.Text("Header shape stroke diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic?.Kind);
            Assert.InRange(warning.LayoutDiagnostic?.X ?? 0D, -1.001D, -0.999D);
            Assert.InRange(warning.LayoutDiagnostic?.Width ?? 0D, 27.999D, 28.001D);
            Assert.True(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterShapeMiterAndShadowUseVisibleBoundsForClippingDiagnostics() {
            OfficeShape shape = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(100, 1),
                new OfficePoint(0, 2));
            shape.StrokeColor = OfficeColor.Black;
            shape.StrokeWidth = 4;
            shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
            shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.5D, 30D, 0D);
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Shape(shape, PdfAlign.Right))
                .Paragraph(paragraph => paragraph.Text("Header miter shadow diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header");
            Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic?.Kind);
            Assert.True((warning.LayoutDiagnostic?.X ?? 0D) + (warning.LayoutDiagnostic?.Width ?? 0D) > options.PageWidth);
            Assert.True(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterUnpaintedAndEmptyClippedShapesDoNotReportVisibleBounds() {
            OfficeShape unpaintedLine = OfficeShape.Line(0, 0, 24, 12);
            unpaintedLine.Transform = OfficeTransform.Translate(300D, 0D);
            OfficeShape emptyClippedShape = OfficeShape.Rectangle(24, 12);
            emptyClippedShape.FillColor = OfficeColor.Blue;
            emptyClippedShape.ClipPath = OfficeClipPath.Empty();
            emptyClippedShape.Shadow = new OfficeShadow(OfficeColor.Black, 0.5D, 300D, 0D);
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header
                    .Shape(unpaintedLine, PdfAlign.Left)
                    .Shape(emptyClippedShape, PdfAlign.Right))
                .Paragraph(paragraph => paragraph.Text("Header invisible shape diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.DoesNotContain(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" ||
                warning.Code == "HeaderFooterContentOverflow" ||
                warning.Code == "HeaderFooterZoneOverlap");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterImageBeyondPhysicalPageReportsClipping() {
            byte[] png = CreateMinimalRgbPng();
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 80,
                MarginRight = 80,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Image(png, 300, 14, PdfAlign.Center))
                .Paragraph(paragraph => paragraph.Text("Physical page clipping diagnostic body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterPageBoundsClipped" &&
                warning.Source == "Header" &&
                warning.LayoutDiagnostic?.Kind == PdfLayoutDiagnosticKind.ClippedContent);
            Assert.True(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterContainedImageUsesVisibleBoundsForClippingDiagnostics() {
            byte[] portraitPng = PdfPngTestImages.CreateRgbPng(1, 4);
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 80,
                MarginRight = 80,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header.Image(portraitPng, 300, 20, PdfAlign.Center, OfficeImageFit.Contain))
                .Paragraph(paragraph => paragraph.Text("Contained header image body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterContentOverflow");
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterZonedContainedImageUsesVisibleBoundsForClippingDiagnostics() {
            byte[] portraitPng = PdfPngTestImages.CreateRgbPng(1, 4);
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 80,
                MarginRight = 80,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header
                    .Zones("H", null, null)
                    .Image(portraitPng, 300, 20, PdfAlign.Left, OfficeImageFit.Contain))
                .Paragraph(paragraph => paragraph.Text("Contained zoned header image body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "HeaderFooterPageBoundsClipped");
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterVisualOnlyGroupsReportOverlap() {
            byte[] png = CreateMinimalRgbPng();
            var report = new PdfConversionReport();
            var options = new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 40,
                MarginRight = 40,
                MarginTop = 50,
                MarginBottom = 50
            }.ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            byte[] bytes = PdfDocument.Create(options)
                .Header(header => header
                    .Image(png, 120, 14, PdfAlign.Left)
                    .Image(png, 120, 14, PdfAlign.Center))
                .Paragraph(paragraph => paragraph.Text("Visual-only header overlap body"))
                .ToBytes();

            Assert.NotEmpty(bytes);
            Assert.Contains(report.Warnings, warning =>
                warning.Code == "HeaderFooterZoneOverlap" &&
                warning.Source == "Header" &&
                warning.Severity == PdfConversionWarningSeverity.Information);
            Assert.False(report.HasLoss);
        }

        [Fact]
        public void HeaderFooterShapes_RenderInsideMarginAreasWithoutImplicitFooterText() {
            OfficeShape headerShape = OfficeShape.Rectangle(20, 10);
            headerShape.FillColor = OfficeColor.Red;
            headerShape.StrokeColor = OfficeColor.Black;
            headerShape.StrokeWidth = 1;

            OfficeShape footerShape = OfficeShape.Rectangle(22, 8);
            footerShape.FillColor = OfficeColor.Blue;
            footerShape.StrokeColor = OfficeColor.Green;
            footerShape.StrokeWidth = 1.5;

            var doc = PdfDocument.Create(new PdfOptions {
                    PageWidth = 300,
                    PageHeight = 200,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 40,
                    MarginBottom = 40,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .Header(header => header.Shape(headerShape, PdfAlign.Center).Text("HeaderShapeText"))
                .Footer(footer => footer.Shape(footerShape, PdfAlign.Right))
                .Paragraph(paragraph => paragraph.Text("Header footer shape body"));

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);

            Assert.Contains("1 0 0 rg", rawPdf);
            Assert.Contains("0 0 1 rg", rawPdf);
            Assert.Contains("0 0.502 0 RG", rawPdf);
            Assert.Contains(" re B", rawPdf);
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            string text = page.Text;
            Assert.Contains("HeaderShapeText", text);
            Assert.Contains("Header footer shape body", text);
            Assert.DoesNotContain("Page 1", text, StringComparison.Ordinal);

            var letters = page.Letters.ToList();
            string letterText = string.Concat(letters.Select(letter => letter.Value));
            int headerStart = letterText.IndexOf("HeaderShapeText", StringComparison.Ordinal);
            Assert.True(headerStart >= 0);
            var headerLetters = letters
                .Skip(headerStart)
                .Take("HeaderShapeText".Length)
                .ToList();
            double textLeft = headerLetters.Min(letter => letter.StartBaseLine.X);
            double textRight = headerLetters.Max(letter => letter.EndBaseLine.X);

            System.Text.RegularExpressions.Match headerPlacement = System.Text.RegularExpressions.Regex.Match(
                rawPdf,
                @"(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) 20 10 re B",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            Assert.True(headerPlacement.Success, "Expected the 20-by-10 header rectangle placement.");
            double headerShapeX = double.Parse(
                headerPlacement.Groups["x"].Value,
                System.Globalization.CultureInfo.InvariantCulture);

            Assert.True(headerShapeX >= textRight + 3.9D, "Aligned header text and shapes must not overlap.");
            Assert.InRange(Math.Abs((textLeft + headerShapeX + 20D) / 2D - 150D), 0D, 0.1D);

            System.Text.RegularExpressions.Match footerPlacement = System.Text.RegularExpressions.Regex.Match(
                rawPdf,
                @"(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) 22 8 re B",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            Assert.True(footerPlacement.Success, "Expected the 22-by-8 footer rectangle placement.");
            double footerShapeX = double.Parse(
                footerPlacement.Groups["x"].Value,
                System.Globalization.CultureInfo.InvariantCulture);
            Assert.InRange(footerShapeX, 247.99D, 248.01D);
        }

    }
}
