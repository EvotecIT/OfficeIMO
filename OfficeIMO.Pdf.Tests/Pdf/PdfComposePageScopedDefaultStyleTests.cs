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
        public void ComposePage_DefaultParagraphStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfParagraphStyle {
                FirstLineIndent = global::OfficeIMO.OfficeLength.Points(24),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            };
            var doc = PdfDocument.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.DefaultParagraphStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageOneFirst").LineBreak().Text("PageOneSecond"))));
                });

                style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(0);

                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageTwoFirst").LineBreak().Text("PageTwoSecond"))));
                });
            });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));

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
                Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDocument.Create();

            doc.Compose(c => {
                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(style);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageOneTextStyle"))));
                });

                style.FontSize = 8;
                style.Color = global::OfficeIMO.OfficeColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.DefaultTextStyle(text => text.Font(PdfStandardFont.Helvetica).FontSize(10));
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("PageTwoTextStyle"))));
                });
            });

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));

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
                Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDocument.Create(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                // The second page inherits this document style, not the page-one override.
                DefaultHeadingStyles = new PdfHeadingStyles {
                    Level2 = new PdfHeadingStyle { Font = PdfStandardFont.Helvetica, FontSize = 18 }
                }
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
                style.Color = global::OfficeIMO.OfficeColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().H2("PageTwoHeading")));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0])).Select(l => l.PointSize).First();
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);

            Assert.InRange(pageOneSize, 12.5, 13.5);
            Assert.InRange(pageTwoSize, 17.5, 18.5);
            Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultListStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfListStyle {
                FontSize = 13,
                LeftIndent = 14,
                Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30)
            };
            var doc = PdfDocument.Create(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                // Keep a distinct document fallback to detect page-style leakage.
                DefaultListStyle = new PdfListStyle { FontSize = 10, LeftIndent = 0, MarkerGap = 2.78 }
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
                style.Color = global::OfficeIMO.OfficeColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().Bullets(new[] { "PageTwoList" })));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

            var page1 = pdf.GetPage(1);
            var page2 = pdf.GetPage(2);
            double pageOneSize = page1.Letters.Where(l => !char.IsWhiteSpace(l.Value[0]) && l.Value != "•").Select(l => l.PointSize).First();
            double pageTwoSize = page2.Letters.Where(l => !char.IsWhiteSpace(l.Value[0]) && l.Value != "•").Select(l => l.PointSize).First();
            double pageOneBulletX = page1.Letters.First(l => l.Value == "•").StartBaseLine.X;
            double pageTwoBulletX = page2.Letters.First(l => l.Value == "•").StartBaseLine.X;
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);

            Assert.InRange(pageOneSize, 12.5, 13.5);
            Assert.InRange(pageTwoSize, 9.5, 10.5);
            Assert.InRange(pageOneBulletX, 43, 45);
            Assert.InRange(pageTwoBulletX, 29.5, 30.5);
            Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultPanelStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfPanelStyle {
                PaddingX = 16,
                MaxWidth = 180,
                Align = PdfAlign.Center,
                Background = global::OfficeIMO.OfficeColor.FromRgb(240, 248, 255)
            };
            var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(360),
                PageHeight = global::OfficeIMO.OfficeLength.Points(240),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
                // Page two measures the document's six-point panel inset.
                DefaultPanelStyle = new PdfPanelStyle { PaddingX = 6, PaddingY = 6 }
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
                style.Background = global::OfficeIMO.OfficeColor.Black;

                c.Page(page => {
                    page.Margin(30);
                    page.Content(content =>
                        content.Column(column =>
                            column.Item().PanelParagraph(p => p.Text("PageTwoPanel"))));
                });
            });

            byte[] bytes = doc.ToBytes();
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

            double pageOneX = FindWordStartX(pdf.GetPage(1), "PageOnePanel");
            double pageTwoX = FindWordStartX(pdf.GetPage(2), "PageTwoPanel");
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);

            Assert.InRange(pageOneX, 105, 107);
            Assert.InRange(pageTwoX, 35, 37);
            Assert.Contains("0.941 0.973 1 rg", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_DefaultHorizontalRuleStyleAppliesOnlyToThatPageAndSnapshotsInput() {
            var style = new PdfHorizontalRuleStyle {
                Thickness = 2,
                Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30),
                SpacingBefore = 3,
                SpacingAfter = 16
            };
            var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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
                style.Color = global::OfficeIMO.OfficeColor.FromRgb(200, 10, 10);
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
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

            double pageOneY = FindWordStartY(pdf.GetPage(1), "PageOneRule");
            double pageTwoY = FindWordStartY(pdf.GetPage(2), "PageTwoRule");
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);

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
            var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);
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
            var doc = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            string rawPdf = PdfTestContentText.ReadDecoded(bytes);
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
                Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30)
            };
            var tableStyle = PdfTableStyles.Minimal();
            tableStyle.CellPaddingX = 22;
            var theme = new PdfTheme {
                TextStyle = textStyle,
                TableStyle = tableStyle
            };
            var doc = PdfDocument.Create(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));

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
            var style = PdfTableStyles.Minimal();
            style.CellPaddingX = 22;
            var doc = PdfDocument.Create(new PdfCreateOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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

            using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));

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
                PdfDocument.Create().Compose(c => c.Page(page => page.DefaultTableStyle("Missing Table Style"))));

            Assert.Equal("styleName", exception.ParamName);
            Assert.Contains("Unsupported Word table style", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ComposePage_ExposesReadOnlyPageBlockCollection() {
            var doc = PdfDocument.Create();

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

    }
}
