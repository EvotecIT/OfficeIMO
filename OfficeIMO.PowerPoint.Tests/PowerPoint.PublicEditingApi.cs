using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPublicEditingApi {
        [Fact]
        public void CreateEditSaveAndReopenUseConcretePresentationObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide first = presentation.AddSlide();
                    first.AddTitle("Canonical title");
                    PowerPointTextBox bullets = first.AddTextBox(string.Empty);
                    bullets.AddBullet("One");
                    bullets.AddBullet("Two");
                    first.AddTable(2, 2);
                    first.Notes.Text = "Notes text";

                    PowerPointSlide second = presentation.AddSlide();
                    second.AddTitle("Second slide");
                    presentation.MoveSlide(1, 0);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    Assert.Equal(2, presentation.Slides.Count);
                    Assert.Equal("Second slide", presentation.Slides[0].TextBoxes.First().Text);
                    Assert.Equal("Notes text", presentation.Slides[1].Notes.Text);
                    Assert.Single(presentation.Slides[1].Tables);
                    presentation.Slides[1].ReplaceText("Canonical", "Edited");
                    presentation.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(filePath,
                    new PowerPointLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });
                Assert.Contains(reopened.Slides[1].TextBoxes, textBox => textBox.Text == "Edited title");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
