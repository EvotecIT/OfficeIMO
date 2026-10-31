using System;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_AddHyperLink_InvalidTextThrows() {
            using var document = WordDocument.Create();
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(null!, new Uri("https://example.com")).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(string.Empty, new Uri("https://example.com")).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(" ", new Uri("https://example.com")).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(null!, "Bookmark").Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(string.Empty, "Bookmark").Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(" ", "Bookmark").Paragraph);
        }

        [Fact]
        public void Test_AddHyperLink_InvalidUriOrAnchorThrows() {
            using var document = WordDocument.Create();
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", (Uri)null!).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", (string)null!).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", string.Empty).Paragraph);
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", " ").Paragraph);
        }
    }
}
