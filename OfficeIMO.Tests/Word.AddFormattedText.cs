using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void AddFormattedTextCreatesRunsWithFormatting() {
            using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "FormattedText.docx"));
            var paragraph = document.AddParagraph("Hello");
            paragraph.AddFormattedText(" bold", bold: true);
            paragraph.AddFormattedText(" italic", italic: true);
            paragraph.AddFormattedText(" underline", underline: UnderlineValues.Single);
            paragraph.AddHyperLink(" link", new Uri("https://example.com"));

            var runs = paragraph.GetRuns().ToList();
            Assert.Equal(5, runs.Count);
            Assert.Equal("Hello", runs[0].Text);
            Assert.True(runs[1].Bold);
            Assert.True(runs[2].Italic);
            Assert.Equal(UnderlineValues.Single, runs[3].Underline);
            Assert.True(runs[4].IsHyperLink);

            document.Save(false);
        }
    }
}
