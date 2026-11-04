using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMetadataTests {
        [Fact]
        public void TitleAndAuthorRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.Properties.Title = "My Diagram";
            document.Properties.Author = "John Doe";
            document.AddPage("Page1");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("My Diagram", loaded.Properties.Title);
            Assert.Equal("John Doe", loaded.Properties.Author);
        }

        [Fact]
        public void NativeInfoSetsMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document
                .ConfigureProperties(info => { info.SetTitle("Fluent Diagram"); info.SetAuthor("Jane Doe"); });
            document.AddPage("Page1");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Fluent Diagram", loaded.Properties.Title);
            Assert.Equal("Jane Doe", loaded.Properties.Author);
        }
    }
}

