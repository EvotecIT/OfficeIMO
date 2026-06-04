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
        public void WrapRichRuns_TreatsTabsAsWordLikeSpacing() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Alpha\tBeta")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "Alpha", "Beta" }, line.ConvertAll(ExtractText).ToArray());
            Assert.False(ExtractLeadingSpace(line[0]));
            Assert.True(ExtractLeadingSpace(line[1]));
            Assert.False(ExtractLeadingSpaceIsExpandable(line[1]));
            Assert.True(ExtractLeadingAdvance(line[1]) > 3);
        }

        [Fact]
        public void WrapRichRuns_PreservesLeadingTabOnEmptyLine() {
            var result = InvokeWrapRichRuns(new[] {
                TextRun.Tab(),
                TextRun.Normal("Indented")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            object segment = Assert.Single(line);
            Assert.Equal("Indented", ExtractText(segment));
            Assert.True(ExtractLeadingSpace(segment));
            Assert.False(ExtractLeadingSpaceIsExpandable(segment));
            Assert.InRange(ExtractLeadingAdvance(segment), 34, 37);
        }

        [Fact]
        public void WrapRichRuns_AdvancesTabsToDefaultHalfInchStops() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A\tB")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.InRange(ExtractLeadingAdvance(line[1]), 27, 29);
        }

        [Fact]
        public void WrapRichRuns_UsesConfiguredDefaultTabStopWidth() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A\tB")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.InRange(ExtractLeadingAdvance(line[1]), 63, 65);
        }

        [Fact]
        public void WrapRichRuns_CarriesDotLeaderFromExplicitTabRun() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots),
                new TextRun("B")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.False(ExtractLeadingSpace(line[0]));
            Assert.True(ExtractLeadingSpace(line[1]));
            Assert.False(ExtractLeadingSpaceIsExpandable(line[1]));
            Assert.Equal(PdfTabLeaderStyle.Dots, ExtractLeadingTabLeader(line[1]));
        }

        [Theory]
        [InlineData(PdfTabLeaderStyle.Hyphens)]
        [InlineData(PdfTabLeaderStyle.Underscores)]
        public void WrapRichRuns_CarriesNonDotLeaderFromExplicitTabRun(PdfTabLeaderStyle leaderStyle) {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(leaderStyle),
                new TextRun("B")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.False(ExtractLeadingSpace(line[0]));
            Assert.True(ExtractLeadingSpace(line[1]));
            Assert.False(ExtractLeadingSpaceIsExpandable(line[1]));
            Assert.Equal(leaderStyle, ExtractLeadingTabLeader(line[1]));
        }

        [Fact]
        public void WrapRichRuns_RightAlignedTabsAccountForFollowingTokenWidth() {
            var shortResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right),
                new TextRun("12")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);
            var longResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right),
                new TextRun("12345")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var shortLine = Assert.Single(ExtractLines(shortResult));
            var longLine = Assert.Single(ExtractLines(longResult));
            double shortAdvance = ExtractLeadingAdvance(shortLine[1]);
            double longAdvance = ExtractLeadingAdvance(longLine[1]);
            double shortWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "12", PdfStandardFont.Helvetica, 12.0);
            double longWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "12345", PdfStandardFont.Helvetica, 12.0);

            Assert.True(longAdvance < shortAdvance, "Expected wider right-aligned tab text to consume less leading advance.");
            Assert.Equal(shortAdvance + shortWidth, longAdvance + longWidth, 1);
        }

        [Fact]
        public void WrapRichRuns_RightAlignedLeadingTabAccountsForFollowingTokenWidth() {
            var shortResult = InvokeWrapRichRuns(new[] {
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right),
                new TextRun("12")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);
            var longResult = InvokeWrapRichRuns(new[] {
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right),
                new TextRun("12345")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var shortLine = Assert.Single(ExtractLines(shortResult));
            var longLine = Assert.Single(ExtractLines(longResult));
            double shortAdvance = ExtractLeadingAdvance(shortLine[0]);
            double longAdvance = ExtractLeadingAdvance(longLine[0]);
            double shortWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "12", PdfStandardFont.Helvetica, 12.0);
            double longWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "12345", PdfStandardFont.Helvetica, 12.0);

            Assert.True(ExtractLeadingSpace(shortLine[0]));
            Assert.True(ExtractLeadingSpace(longLine[0]));
            Assert.True(longAdvance < shortAdvance, "Expected wider leading right-aligned tab text to consume less leading advance.");
            Assert.Equal(shortAdvance + shortWidth, longAdvance + longWidth, 1);
        }

        [Fact]
        public void WrapRichRuns_CenterAlignedTabsCenterFollowingTokenOnStop() {
            var shortResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Center),
                new TextRun("AB")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);
            var longResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Center),
                new TextRun("ABCDE")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var shortLine = Assert.Single(ExtractLines(shortResult));
            var longLine = Assert.Single(ExtractLines(longResult));
            double shortAdvance = ExtractLeadingAdvance(shortLine[1]);
            double longAdvance = ExtractLeadingAdvance(longLine[1]);
            double shortWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "AB", PdfStandardFont.Helvetica, 12.0);
            double longWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "ABCDE", PdfStandardFont.Helvetica, 12.0);

            Assert.True(longAdvance < shortAdvance, "Expected wider center-aligned tab text to consume less leading advance.");
            Assert.Equal(shortAdvance + shortWidth / 2D, longAdvance + longWidth / 2D, 1);
        }

        [Fact]
        public void WrapRichRuns_DecimalAlignedTabsAlignDecimalSeparator() {
            var shortResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator),
                new TextRun("12.30")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);
            var longResult = InvokeWrapRichRuns(new[] {
                new TextRun("A"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.DecimalSeparator),
                new TextRun("1234.50")
            }, 200, 12, PdfStandardFont.Helvetica, tabStopWidth: 72);

            var shortLine = Assert.Single(ExtractLines(shortResult));
            var longLine = Assert.Single(ExtractLines(longResult));
            double shortAdvance = ExtractLeadingAdvance(shortLine[1]);
            double longAdvance = ExtractLeadingAdvance(longLine[1]);
            double shortPrefixWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "12", PdfStandardFont.Helvetica, 12.0);
            double longPrefixWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "1234", PdfStandardFont.Helvetica, 12.0);

            Assert.True(longAdvance < shortAdvance, "Expected wider decimal prefix text to consume less leading advance.");
            Assert.Equal(shortAdvance + shortPrefixWidth, longAdvance + longPrefixWidth, 1);
        }
    }
}
