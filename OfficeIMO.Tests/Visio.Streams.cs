using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStreamTests {
        [Fact]
        public void Create_Save_Load_Stream_Roundtrip() {
            using var stream = new MemoryStream();
            VisioDocument document = VisioDocument.Create(stream);
            document.AddPage("StreamPage");
            document.Save();

            Assert.True(stream.Length > 0);
            stream.Position = 0;

            VisioDocument loaded = VisioDocument.Load(stream);
            Assert.Single(loaded.Pages);
            Assert.Equal("StreamPage", loaded.Pages[0].Name);
        }
    }
}
