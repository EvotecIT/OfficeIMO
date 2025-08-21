using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMetadataTests {
        [Fact]
        public void TitleAndAuthorRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.Title = "My Diagram";
            document.Author = "John Doe";
            document.AddPage("Page1");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("My Diagram", loaded.Title);
            Assert.Equal("John Doe", loaded.Author);
        }

        [Fact]
        public void FluentInfoSetsMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .Info(info => info.Title("Fluent Diagram").Author("Jane Doe"))
                .End();
            document.AddPage("Page1");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("Fluent Diagram", loaded.Title);
            Assert.Equal("Jane Doe", loaded.Author);
        }
    }
}

