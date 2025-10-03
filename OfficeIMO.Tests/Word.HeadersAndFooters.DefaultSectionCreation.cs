using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void HeaderAndFooterAccessors_CreateSectionWhenMissing() {
            using var stream = new MemoryStream();
            using var document = WordDocument.Create(stream);

            document.Sections.Clear();

            var header = document.HeaderDefaultOrCreate;
            Assert.NotNull(header);
            Assert.Single(document.Sections);

            var footer = document.FooterDefaultOrCreate;
            Assert.NotNull(footer);
            Assert.Single(document.Sections);

            Assert.False(document.DifferentFirstPage);
            Assert.Single(document.Sections);

            document.DifferentFirstPage = true;
            Assert.True(document.DifferentFirstPage);

            Assert.False(document.DifferentOddAndEvenPages);

            document.DifferentOddAndEvenPages = true;
            Assert.True(document.DifferentOddAndEvenPages);
        }

        [Fact]
        public void HeaderAndFooterProperties_ThrowWhenSectionMissing() {
            using var stream = new MemoryStream();
            using var document = WordDocument.Create(stream);

            document.Sections.Clear();

            Assert.Throws<InvalidOperationException>(() => _ = document.Header);
            Assert.Throws<InvalidOperationException>(() => _ = document.Footer);
        }
    }
}
