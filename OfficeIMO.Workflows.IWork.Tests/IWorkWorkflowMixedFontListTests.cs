using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    private static void AssertMixedFontList(PdfReadDocument pdf, bool edited, string title = "Mixed family decimal items") {
        var page = Assert.Single(pdf.Pages); string text = Compact(page.ExtractText());
        var expected = new List<string> { "7." + (edited ? "Root native" : "Root seven") };
        for (int level = 1; level <= 8; level++)
            expected.Add("1." + (edited && level == 8 ? "Deep native" : "Depth " + level));
        expected.AddRange(["2.Deep two", "8.Root eight", "1.Child one", "(3)Serif marker", "(1)Serif child",
            "5)Mono marker", "1)Mono child", "2.Resumed child"]);
        int previous = -1;
        foreach (string item in expected) {
            string wanted = Compact(item + " serif mono plain");
            int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
            Assert.True(at > previous, "Missing or reordered mixed-font item: " + wanted); previous = at;
        }
        if (edited) Assert.True(text.IndexOf("3.Nativechildplain", StringComparison.Ordinal) > previous);
        Assert.Contains(Compact(title), text); Assert.Contains("Plaingapretainscounters", text);
        var spans = page.GetTextSpans();
        Assert.Contains(spans, span => span.BaseFont == "Times-Bold" && span.IsBold && !span.IsItalic
            && span.FontSize == 12 && span.Color?.ToString() == "#003366");
        Assert.Contains(spans, span => span.BaseFont == "Courier-Oblique" && !span.IsBold && span.IsItalic
            && span.FontSize == 12 && span.Color?.ToString() == "#008000");
        Assert.Contains(spans, span => span.Text == "plain" && span.BaseFont == "Helvetica" && !span.IsBold && !span.IsItalic);
        Assert.All(spans.Where(span => Regex.IsMatch(span.Text.Trim(), @"^(\d+\.|\(\d+\)|\d+\))$")), span => {
            Assert.False(span.IsBold); Assert.False(span.IsItalic);
        });
        Assert.Contains(spans, span => span.Text == "(3)" && span.BaseFont == "Times-Roman");
        Assert.Contains(spans, span => span.Text == "5)" && span.BaseFont == "Courier");

        static string Compact(string value) => Regex.Replace(value, @"\s+", string.Empty);
    }
}
