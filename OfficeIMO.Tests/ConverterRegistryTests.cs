using System.IO;
using System.Text;
using OfficeIMO.Converters;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConverterRegistryTests {
        [Fact]
        public void RegisteredConverterCanBeResolved() {
            ConverterRegistry.Register("markdown->word-test", () => new MarkdownToWordConverter());
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes("# Title"));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = ConverterRegistry.Resolve("markdown->word-test");
            converter.Convert(input, output, new MarkdownToWordOptions());
            Assert.True(output.Length > 0);
        }

        [Fact]
        public void ResolvingMissingConverterThrows() {
            Assert.Throws<System.InvalidOperationException>(() => ConverterRegistry.Resolve("missing"));
        }
    }
}
