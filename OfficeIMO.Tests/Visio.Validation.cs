using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioValidation {
        [Fact]
        public void SavedDocumentValidatesAndOpens() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Start"));
            document.Save();

            var issues = VisioValidator.Validate(filePath);
            Assert.Empty(issues);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages);
        }
    }
}

