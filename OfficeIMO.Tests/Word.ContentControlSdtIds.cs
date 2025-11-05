using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ContentControls_AssignUniqueIds() {
            using var document = WordDocument.Create();
            var primaryParagraph = document.AddParagraph();

            var imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            primaryParagraph.AddStructuredDocumentTag("Structured");
            primaryParagraph.AddCheckBox();
            primaryParagraph.AddDatePicker(DateTime.Today);
            primaryParagraph.AddDropDownList(new[] { "One", "Two" });
            primaryParagraph.AddComboBox(new[] { "Red", "Blue" }, defaultValue: "Red");

            var pictureParagraph = document.AddParagraph();
            pictureParagraph.AddPictureControl(imagePath, 50, 50);

            var repeatingParagraph = document.AddParagraph();
            repeatingParagraph.AddRepeatingSection(sectionTitle: "Items");

            var ids = document._document
                .Descendants<SdtId>()
                .Select(id => id.Val?.Value)
                .Where(val => val.HasValue)
                .Select(val => val!.Value)
                .ToArray();

            Assert.Equal(7, ids.Length);
            Assert.Equal(ids.Length, ids.Distinct().Count());
            Assert.All(ids, id => Assert.InRange(id, 1, int.MaxValue));
        }

        [Fact]
        public void ContentControls_PreserveUniquenessAcrossReload() {
            string filePath = Path.Combine(_directoryWithFiles, "UniqueSdtIds.docx");

            int initialMaxId;
            using (var document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Initial");
                paragraph.AddStructuredDocumentTag("First");
                paragraph.AddCheckBox();

                document.Save(false);

                initialMaxId = document._document
                    .Descendants<SdtId>()
                    .Select(id => id.Val?.Value ?? 0)
                    .DefaultIfEmpty(0)
                    .Max();
            }

            using (var document = WordDocument.Load(filePath)) {
                var paragraph = document.AddParagraph("Next");
                paragraph.AddDatePicker();
                paragraph.AddDropDownList(new[] { "A", "B" });
                paragraph.AddComboBox(new[] { "Red", "Blue" }, defaultValue: "Blue");

                document.Save(false);

                var allIds = document._document
                    .Descendants<SdtId>()
                    .Select(id => id.Val?.Value ?? 0)
                    .Where(id => id > 0)
                    .ToArray();

                Assert.Equal(allIds.Length, allIds.Distinct().Count());
                Assert.All(allIds, id => Assert.InRange(id, 1, int.MaxValue));

                var newlyAllocated = allIds.Where(id => id > initialMaxId).ToArray();
                Assert.True(newlyAllocated.Length >= 3, "Expected the reloaded document to allocate new identifiers.");
                Assert.All(newlyAllocated, id => Assert.True(id > initialMaxId));
            }
        }
    }
}
