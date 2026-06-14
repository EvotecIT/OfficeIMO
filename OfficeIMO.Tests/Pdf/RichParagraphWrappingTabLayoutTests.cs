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
        public void WrapRichRuns_UsesExplicitParagraphTabStopsBeforeDefaultTabWidth() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A\tB\tC")
            }, 240, 12, PdfStandardFont.Helvetica, tabStopWidth: 36, tabStops: new[] {
                new PdfTabStop(90),
                new PdfTabStop(150)
            });

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B", "C" }, line.ConvertAll(ExtractText).ToArray());
            Assert.InRange(ExtractLeadingAdvance(line[1]), 81, 83);
            Assert.InRange(ExtractLeadingAdvance(line[2]), 51, 53);
        }

        [Fact]
        public void WrapRichRuns_OffsetsExplicitTabStopsForFirstLineIndent() {
            var method = typeof(PdfWriter).GetMethod("WrapRichRunsCoreWithFirstLineOrigin", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);

            var result = method!.Invoke(null, new object?[] {
                new[] { new TextRun("A\tB") },
                180D,
                12D,
                PdfStandardFont.Helvetica,
                16.8D,
                156D,
                24D,
                36D,
                null,
                new[] { new PdfTabStop(90) }
            })!;

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.InRange(ExtractLeadingAdvance(line[1]), 57, 59);
        }

        [Fact]
        public void WrapRichRuns_UsesExplicitRightAlignedTabStopAndLeader() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Revenue"),
                TextRun.Tab(),
                new TextRun("123")
            }, 240, 12, PdfStandardFont.Helvetica, tabStops: new[] {
                new PdfTabStop(180, PdfTabAlignment.Right, PdfTabLeaderStyle.Dots)
            });

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "Revenue", "123" }, line.ConvertAll(ExtractText).ToArray());
            double revenueWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "Revenue", PdfStandardFont.Helvetica, 12.0);
            double valueWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "123", PdfStandardFont.Helvetica, 12.0);

            Assert.Equal(PdfTabLeaderStyle.Dots, ExtractLeadingTabLeader(line[1]));
            Assert.Equal(180D, revenueWidth + ExtractLeadingAdvance(line[1]) + valueWidth, 1);
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
        public void WrapRichRuns_RightAlignedTabsClampOversizedStopsToTextFrame() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Column evidence"),
                TextRun.Tab(PdfTabLeaderStyle.Dots, PdfTabAlignment.Right),
                new TextRun("1")
            }, 180, 12, PdfStandardFont.Helvetica, tabStopWidth: 432);

            var line = Assert.Single(ExtractLines(result));
            double usedWidth = line.Sum(segment =>
                ExtractLeadingAdvance(segment) +
                InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", ExtractText(segment), PdfStandardFont.Helvetica, 12.0));

            Assert.Equal(new[] { "Column", "evidence", "1" }, line.ConvertAll(ExtractText).ToArray());
            Assert.Equal(PdfTabLeaderStyle.Dots, ExtractLeadingTabLeader(line[2]));
            Assert.InRange(usedWidth, 178, 180.5);
        }

        [Fact]
        public void CalculateTabAdvance_BoundedRightTabCanShrinkBelowSpace() {
            double advance = InvokePrivateFontMethod<double>(
                "CalculateTabAdvance",
                10D,
                89D,
                3D,
                PdfTabAlignment.Right,
                36D,
                string.Empty,
                PdfStandardFont.Helvetica,
                12D,
                PdfTextBaseline.Normal,
                null!,
                100D);

            Assert.Equal(1D, advance);
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
