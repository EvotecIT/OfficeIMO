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
        public void LogicalEndAlignmentSavesAsRightAlignedHorzAlign() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Notes");
            VisioShape textBox = page.AddTextBox("note", OfficeLength.Inches(4), OfficeLength.Inches(7), OfficeLength.Inches(5), OfficeLength.Inches(0.6), "Aligned");
            textBox.TextStyle = new VisioTextStyle { HorizontalAlignment = OfficeAlignment.End };

            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(OfficeAlignment.Right, Assert.Single(loaded.Pages[0].Shapes).TextStyle?.HorizontalAlignment);
        }

        [Fact]
        public void PageCanAddEditableTextBoxAdornment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Notes");
            VisioShape textBox = page.AddTextBox("title", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.6, page.DefaultUnit)), "Generated diagram title");
            textBox.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 18,
                Bold = true,
                HorizontalAlignment = OfficeAlignment.Center
            };

            Assert.Equal("Text Box", textBox.NameU);
            Assert.Equal("title", textBox.Id);
            Assert.Equal(0, textBox.LinePattern);
            Assert.Equal(0, textBox.FillPattern);
            Assert.NotNull(textBox.FindUserCell("OfficeIMO.Kind"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedTextBox = Assert.Single(loaded.Pages[0].Shapes);
            Assert.Equal("title", loadedTextBox.Id);
            Assert.Equal("Generated diagram title", loadedTextBox.Text);
            Assert.Equal("Text Box", loadedTextBox.NameU);
        }
    }
}

