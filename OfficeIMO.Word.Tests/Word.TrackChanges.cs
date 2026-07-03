using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TrackChangesToggle() {
            string filePath = Path.Combine(_directoryWithFiles, "Test_TrackChanges.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.TrackChanges = true;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.TrackChanges);
                Assert.True(document.Settings.TrackFormatting);
                Assert.True(document.Settings.TrackMoves);
                document.TrackChanges = false;
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.False(document.TrackChanges);
                Assert.False(document.Settings.TrackFormatting);
                Assert.False(document.Settings.TrackMoves);
            }
        }
    }
}
