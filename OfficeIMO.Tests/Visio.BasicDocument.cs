using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioBasicDocument {
        [Fact]
        public void CanCreateBasicVisioDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page1");
            VisioShape shape1 = new("S1");
            VisioShape shape2 = new("S2");
            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);
            page.Connectors.Add(new VisioConnector(shape1, shape2));

            Assert.Single(document.Pages);
            Assert.Equal(2, page.Shapes.Count);
            Assert.Single(page.Connectors);
            document.Save();
        }

        [Fact]
        public void PageCanAddEditableTextBoxAdornment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Notes");
            VisioShape textBox = page.AddTextBox(4, 7, 5, 0.6, "Generated diagram title");
            textBox.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 18,
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center
            };

            Assert.Equal("Text Box", textBox.NameU);
            Assert.Equal(0, textBox.LinePattern);
            Assert.Equal(0, textBox.FillPattern);
            Assert.NotNull(textBox.FindUserCell("OfficeIMO.Kind"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedTextBox = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal("Generated diagram title", loadedTextBox.Text);
            Assert.Equal("Text Box", loadedTextBox.NameU);
        }
    }
}

