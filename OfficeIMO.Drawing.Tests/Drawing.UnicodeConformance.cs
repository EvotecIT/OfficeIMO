using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Samples of the Unicode 16.0.0 conformance files for UAX #9 and UAX #14.</summary>
public class DrawingUnicodeConformanceTests {
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Unicode", name);

    private static IEnumerable<string> DataLines(string name) =>
        File.ReadLines(Fixture(name)).Where(static line => line.Length > 0 && line[0] != '#');

    [Fact]
    public void BidiLevelsAndVisualOrderMatchBidiCharacterTest() {
        int cases = 0;
        foreach (string line in DataLines("BidiCharacterTest.sample.txt")) {
            string[] fields = line.Split(';');
            int[] scalars = fields[0].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(static hex => Convert.ToInt32(hex, 16)).ToArray();
            int direction = int.Parse(fields[1]);
            string[] levels = fields[3].Split(' ', StringSplitOptions.RemoveEmptyEntries);
            int[] order = fields[4].Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToArray();

            OfficeUnicodeBidi.Paragraph paragraph = OfficeUnicodeBidi.Resolve(scalars, direction == 2 ? null : direction);
            Assert.Equal(int.Parse(fields[2]), paragraph.ParagraphLevel);
            byte[] resolved = paragraph.GetLineLevels(0, scalars.Length);
            for (int index = 0; index < scalars.Length; index++) {
                if (levels[index] != "x") Assert.True(levels[index] == resolved[index].ToString(), line);
            }
            int[] kept = Enumerable.Range(0, scalars.Length).Where(index => levels[index] != "x").ToArray();
            int[] visual = OfficeUnicodeBidi.ReorderLine(kept.Select(index => resolved[index]).ToArray()).Select(index => kept[index]).ToArray();
            Assert.True(order.SequenceEqual(visual), line);
            cases++;
        }
        Assert.True(cases > 300);
    }

    [Fact]
    public void LineBreakOpportunitiesMatchLineBreakTest() {
        int cases = 0;
        foreach (string line in DataLines("LineBreakTest.sample.txt")) {
            var scalars = new List<int>();
            var breaks = new List<bool>();
            foreach (string token in line.Split(' ', StringSplitOptions.RemoveEmptyEntries)) {
                if (token == "×") breaks.Add(false);
                else if (token == "÷") breaks.Add(true);
                else scalars.Add(Convert.ToInt32(token, 16));
            }
            OfficeUnicodeLineBreaker.BreakAction[] actions = OfficeUnicodeLineBreaker.Analyze(scalars.ToArray());
            for (int index = 1; index < scalars.Count; index++) {
                Assert.True(breaks[index] == (actions[index] != OfficeUnicodeLineBreaker.BreakAction.Prohibited), line);
            }
            cases++;
        }
        Assert.True(cases > 400);
    }

    [Theory]
    [InlineData("state-of-the-art", new[] { 6, 9, 13 })]
    [InlineData("a/b", new[] { 2 })]
    [InlineData("a​b", new[] { 2 })]
    [InlineData("a b", new int[0])]
    [InlineData("1,000.50", new int[0])]
    [InlineData("你好。世界", new[] { 1, 3, 4 })]
    public void PublicBreakPositionsFollowUax14(string text, int[] expected) =>
        Assert.Equal(expected, OfficeTextLineBreaks.GetBreakPositions(text));
}
