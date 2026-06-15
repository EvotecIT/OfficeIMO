using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
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

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Cover body."))
                .PageBreak()
                .Paragraph(p => p.Text("Running body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Cover body."))
                .PageBreak()
                .Paragraph(p => p.Text("Running body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.DoesNotContain("Runningheader", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Runningfooter", page1Text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Runningheader", page2Text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Runningfooter", page2Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ClearingFirstPageWatermarkDoesNotEnableBlankHeaderFooterVariant() {
        var options = new PdfOptions {
            ShowHeader = true,
            HeaderFormat = "Running header"
        };

        byte[] bytes = PdfDocument.Create(options)
            .Page(page => {
                page.FirstPageWatermark((PdfTextWatermark?)null);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Body text."))));
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string pageText = Normalize(pdf.GetPage(1).Text);

        Assert.Contains("Runningheader", pageText, StringComparison.OrdinalIgnoreCase);
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

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(p => p.Text("First body."))
                .PageBreak()
                .Paragraph(p => p.Text("Even body."))
                .PageBreak()
                .Paragraph(p => p.Text("Odd body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
        public void DifferentFirstAndEvenPageWatermarks_UseMatchingPageVariant() {
            var options = new PdfOptions {
                DifferentFirstPageHeaderFooter = true,
                DifferentOddAndEvenPagesHeaderFooter = true,
                TextWatermark = new PdfTextWatermark("Odd watermark") { Opacity = 0.18 },
                FirstPageTextWatermark = new PdfTextWatermark("First watermark") { Opacity = 0.18 },
                EvenPageTextWatermark = new PdfTextWatermark("Even watermark") { Opacity = 0.18 }
            };

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Page one body."))
                .PageBreak()
                .Paragraph(p => p.Text("Page two body."))
                .PageBreak()
                .Paragraph(p => p.Text("Page three body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            Assert.Equal(3, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);
            string page3Text = Normalize(pdf.GetPage(3).Text);

            Assert.Contains("Firstwatermark", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evenwatermark", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddwatermark", page1Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Evenwatermark", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Firstwatermark", page2Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Oddwatermark", page2Text, StringComparison.OrdinalIgnoreCase);

            Assert.Contains("Oddwatermark", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Firstwatermark", page3Text, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Evenwatermark", page3Text, StringComparison.OrdinalIgnoreCase);
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

            byte[] bytes = PdfDocument.Create(options)
                .Paragraph(p => p.Text("Odd body."))
                .PageBreak()
                .Paragraph(p => p.Text("Even body."))
                .PageBreak()
                .Paragraph(p => p.Text("Odd again body."))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create()
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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
            Assert.Equal(2, pdf.NumberOfPages);

            string page1Text = Normalize(pdf.GetPage(1).Text);
            string page2Text = Normalize(pdf.GetPage(2).Text);

            Assert.Contains("Doc5/6", page1Text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Doc6/6", page2Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterPageTokens_UseConfiguredPageNumberStyleForFormatsAndSegments() {
            var doc = PdfDocument.Create()
                .PageNumberStyle(PdfPageNumberStyle.UpperRoman)
                .Header(header => header.Text("Roman header {page}/{pages}"))
                .Footer(footer => footer.Text(text => text.Text("Roman footer ").CurrentPage().Text("/").TotalPages()))
                .Paragraph(p => p.Text("Roman first body."))
                .PageBreak()
                .Paragraph(p => p.Text("Roman second body."));

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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
            var doc = PdfDocument.Create();

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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
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

    }
}
