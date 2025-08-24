using System;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        [Fact]
        public void Test_AddHyperLink_InvalidTextThrows() {
            using var document = WordDocument.Create();
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(null!, new Uri("https://example.com")));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(string.Empty, new Uri("https://example.com")));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(" ", new Uri("https://example.com")));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(null!, "Bookmark"));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(string.Empty, "Bookmark"));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink(" ", "Bookmark"));
        }

        [Fact]
        public void Test_AddHyperLink_InvalidUriOrAnchorThrows() {
            using var document = WordDocument.Create();
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", (Uri)null!));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", (string)null!));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", string.Empty));
            Assert.Throws<ArgumentException>(() => document.AddHyperLink("Test", " "));
        }
    }
}
