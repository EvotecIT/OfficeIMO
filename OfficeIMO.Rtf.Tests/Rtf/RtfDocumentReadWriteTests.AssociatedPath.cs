using OfficeIMO.Rtf;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public async Task AssociatedPathChangesOnlyWithFileOperationsAndDisposalDoesNotSave() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-rtf-lifecycle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.rtf");
            using (RtfDocument draft = RtfDocument.Create(source)) {
                draft.AddParagraph().AddText("Unsaved draft");
                Assert.False(File.Exists(source));
            }
            Assert.False(File.Exists(source));
            using (RtfDocument document = RtfDocument.Create(source)) {
                document.AddParagraph().AddText("Saved source");
                document.Save();
            }
            byte[] original = File.ReadAllBytes(source);
            using RtfDocument loaded = await RtfDocument.LoadAsync(source);
            Assert.Equal(source, loaded.FilePath);
            loaded.AddParagraph().AddText("Saved copy");
            string copy = Path.Combine(root, "copy.rtf");
            await loaded.SaveAsync(copy);
            Assert.Equal(copy, loaded.FilePath);
            Assert.Equal(original, File.ReadAllBytes(source));
            using var stream = new MemoryStream();
            loaded.Save(stream);
            Assert.True(stream.CanWrite);
            Assert.Equal(copy, loaded.FilePath);
            loaded.AddParagraph().AddText("Associated save");
            await loaded.SaveAsync();
            using RtfDocument reopened = RtfDocument.Load(copy);
            Assert.Equal(copy, reopened.FilePath);
            Assert.Contains(reopened.Paragraphs, paragraph => paragraph.ToPlainText() == "Associated save");
            using RtfDocument detached = RtfDocument.Parse(@"{\rtf1\ansi Detached\par}");
            Assert.Null(detached.FilePath);
            Assert.Throws<InvalidOperationException>(() => detached.Save());
        } finally {
            Directory.Delete(root, true);
        }
    }
}
