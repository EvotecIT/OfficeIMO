using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void ParagraphTabs_RenderAsVisibleDefaultTabStopGap() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    DefaultFontSize = 12
                })
                .Paragraph(p => p.Text("A B"), style: new PdfParagraphStyle {
                    SpacingAfter = 0
                })
                .Paragraph(p => p.Text("A\tB"), style: new PdfParagraphStyle {
                    SpacingAfter = 0
                })
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var lineGaps = page.Letters
                .Where(letter => letter.Value == "A" || letter.Value == "B")
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => {
                    var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
                    var a = ordered.First(letter => letter.Value == "A");
                    var b = ordered.First(letter => letter.Value == "B");
                    return b.StartBaseLine.X - a.EndBaseLine.X;
                })
                .ToArray();

            Assert.Equal(2, lineGaps.Length);
            Assert.True(lineGaps[1] > lineGaps[0] + 10, $"Expected a default tab-stop gap rather than a collapsed single space. Plain gap: {lineGaps[0]:0.##}, tab gap: {lineGaps[1]:0.##}.");
        }

        [Fact]
        public void ParagraphTabs_UseDefaultParagraphStyleTabStopWidth() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 72,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("A\tB"))
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var letters = pdf.GetPage(1).Letters
                .Where(letter => letter.Value == "A" || letter.Value == "B")
                .OrderBy(letter => letter.StartBaseLine.X)
                .ToList();
            var a = Assert.Single(letters, letter => letter.Value == "A");
            var b = Assert.Single(letters, letter => letter.Value == "B");
            double gap = b.StartBaseLine.X - a.EndBaseLine.X;

            Assert.True(gap > 50, $"Expected custom 72pt tab stop to create a wide visible gap. Gap: {gap:0.##}.");
        }

        [Fact]
        public void ParagraphTabs_RenderDotLeadersAndStructuredLeaderReadback() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("Revenue").Tab(PdfTabLeaderStyle.Dots).Text("12"))
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            int dotCount = page.Letters.Count(letter => letter.Value == ".");
            Assert.True(dotCount >= 3, $"Expected dotted leaders to render between the label and value. Dot count: {dotCount}.");

            var structuredPage = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }));
            var leader = Assert.Single(structuredPage.LeaderRows);
            Assert.Equal(new[] { "Revenue", "12" }, leader);
        }

        [Theory]
        [InlineData(PdfTabLeaderStyle.Hyphens, "-")]
        [InlineData(PdfTabLeaderStyle.Underscores, "_")]
        public void ParagraphTabs_RenderNonDotLeaders(PdfTabLeaderStyle leaderStyle, string expectedGlyph) {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("Status").Tab(leaderStyle).Text("Ready"))
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var glyphCount = pdf.GetPage(1).Letters.Count(letter => letter.Value == expectedGlyph);

            Assert.True(glyphCount >= 3, $"Expected {leaderStyle} tab leaders to render with '{expectedGlyph}' glyphs. Glyph count: {glyphCount}.");
        }

        [Fact]
        public void ParagraphTabs_RightAlignedDotLeadersAlignValueEnds() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 200,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("A").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("12"))
                .Paragraph(p => p.Text("Longer").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("12345"))
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var digitEnds = page.Letters
                .Where(letter => char.IsDigit(letter.Value[0]))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => group.Max(letter => letter.EndBaseLine.X))
                .ToArray();

            Assert.Equal(2, digitEnds.Length);
            Assert.InRange(Math.Abs(digitEnds[0] - digitEnds[1]), 0, 1.5);
        }

        [Fact]
        public void ParagraphTabs_DecimalAlignedDotLeadersAlignDecimalSeparators() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 200,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("Tax").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("8.50"))
                .Paragraph(p => p.Text("Total").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("12845.75"))
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var decimalStarts = page.Letters
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => {
                    var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
                    for (int i = 1; i < ordered.Count - 1; i++) {
                        if (ordered[i].Value == "." &&
                            char.IsDigit(ordered[i - 1].Value[0]) &&
                            char.IsDigit(ordered[i + 1].Value[0])) {
                            return ordered[i].StartBaseLine.X;
                        }
                    }

                    throw new InvalidOperationException("Could not find a decimal separator surrounded by digits.");
                })
                .ToArray();

            Assert.Equal(2, decimalStarts.Length);
            Assert.InRange(Math.Abs(decimalStarts[0] - decimalStarts[1]), 0, 1.5);
        }

        [Fact]
        public void ParagraphTabs_DotLeaderReadbackPreservesDecimalValues() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 220,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("Tax").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("8.50"))
                .Paragraph(p => p.Text("Total").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator).Text("1450.75"))
                .Paragraph(p => p.Text("Discount").Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right).Text("$1,234.50"))
                .ToBytes();

            var structuredPage = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }));

            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Tax" && row[1] == "8.50");
            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Total" && row[1] == "1450.75");
            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Discount" && row[1] == "$1,234.50");
        }

        [Fact]
        public void ParagraphTabs_NonDotLeaderReadbackPreservesNumericValues() {
            byte[] bytes = PdfDocument.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 220,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 12,
                    DefaultParagraphStyle = new PdfParagraphStyle {
                        DefaultTabStopWidth = 216,
                        SpacingAfter = 0
                    }
                })
                .Paragraph(p => p.Text("Milestone").Tab(PdfTabLeaderStyle.Hyphens, PdfTabAlignment.Right).Text("Q4"))
                .Paragraph(p => p.Text("Signature count").Tab(PdfTabLeaderStyle.Underscores, PdfTabAlignment.Right).Text("12"))
                .Paragraph(p => p.Text("Balance").Tab(PdfTabLeaderStyle.Hyphens, PdfTabAlignment.DecimalSeparator).Text("$1,234.50"))
                .ToBytes();

            var structuredPage = Assert.Single(PdfTextExtractor.ExtractStructuredByPage(bytes, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            }));

            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Milestone" && row[1] == "Q4");
            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Signature count" && row[1] == "12");
            Assert.Contains(structuredPage.LeaderRows, row => row.Length >= 2 && row[0] == "Balance" && row[1] == "$1,234.50");
        }
    }
}
