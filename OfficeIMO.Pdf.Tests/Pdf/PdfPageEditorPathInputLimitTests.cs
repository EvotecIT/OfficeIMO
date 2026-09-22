using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageEditorPathInputLimitTests {
    [Fact]
    public void PageEditorPathRoutesRejectOversizedInputBeforeBufferingOrWriting() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-page-editor-limit-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(root, "input.pdf");
        string outputPath = Path.Combine(root, "output.pdf");
        try {
            Directory.CreateDirectory(root);
            using (var file = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write)) {
                file.SetLength(PdfLoadOptions.Default.Limits.MaxInputBytes + 1);
            }

            Action<string>[] routes = {
                input => PdfPageEditor.DeletePages(input, 1),
                input => PdfPageEditor.DuplicatePages(input, 1),
                input => PdfPageEditor.MovePages(input, 1, 1),
                input => PdfPageEditor.ReorderPages(input, 1),
                input => PdfPageEditor.RotatePages(input, 90, 1),
                input => PdfPageEditor.ResizePages(input, new PageSize(612, 792), 1),
                input => PdfPageEditor.SetPageBox(input, "TrimBox", 0, 0, 612, 792, 1)
            };
            foreach (Action<string> route in routes) {
                Assert.Equal(PdfReadLimitKind.InputBytes,
                    Assert.Throws<PdfReadLimitException>(() => route(inputPath)).Kind);
            }

            Assert.Equal(PdfReadLimitKind.InputBytes,
                Assert.Throws<PdfReadLimitException>(() => PdfPageEditor.ReversePages(inputPath, outputPath)).Kind);
            Assert.False(File.Exists(outputPath));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }
}
