using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordDocument_PaintsFullParagraphWhenResolvedFontNeedsMoreThanOneLine(bool styledRuns) {
        const string caption = "Observed events per minute across six intervals: 42, 58, 73, 91, 67 and 82. Capacity is 100 in every interval.";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try {
            using (var document = WordDocument.Create(path)) {
                document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
                document.Margins.Type = WordMargin.Normal;
                if (styledRuns) {
                    var first = document.AddParagraph().AddText("Observed events per minute across six intervals: 42, 58, 73, 91, 67 and 82. ");
                    first.FontFamily = "Arial";
                    first.FontSizePoints = 11D;
                    var second = first.AddText("Capacity is 100 in every interval.");
                    second.FontFamily = "Arial";
                    second.FontSizePoints = 11D;
                    second.Bold = true;
                } else {
                    var paragraph = document.AddParagraph().AddText(caption);
                    paragraph.FontFamily = "Arial";
                    paragraph.FontSizePoints = 11D;
                }
                document.Save();
            }

            using var loaded = WordDocument.Load(path);
            var result = loaded.ExportImage(OfficeImageExportFormat.Svg);
            XNamespace ns = "http://www.w3.org/2000/svg";
            var svg = XDocument.Parse(Encoding.UTF8.GetString(result.Bytes));
            var text = svg.Descendants(ns + "text").Select(item => item.Value).ToArray();
            string painted = Regex.Replace(string.Join(" ", text), @"\s+", " ").Trim();
            Assert.Equal(caption, painted);
            Assert.DoesNotContain("...", painted, StringComparison.Ordinal);
            Assert.DoesNotContain("\u2026", painted, StringComparison.Ordinal);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }
}
