using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void PageEditor_RejectsInvalidSelections() {
        byte[] source = BuildThreePagePdf();

        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, 1, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, 1, 2, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePages(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePages(source, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages((Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages(source, null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages(new MemoryStream(source), null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(new MemoryStream(source), new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePages("input.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages("missing.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(" ", new MemoryStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages(" ", "out.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePages("missing.pdf", " ", 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 3, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(source, 1, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRange(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRange("input.pdf", (Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange("missing.pdf", new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(" ", new MemoryStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRange("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, PdfPageRange.From(1, 3)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DeletePageRanges(source, PdfPageRange.From(1, 2), PdfPageRange.From(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(source, null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(source, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DeletePageRanges(new MemoryStream(source), null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(new MemoryStream(source), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges(" ", "out.pdf", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DeletePageRanges("missing.pdf", " ", PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePages(source, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePages(source, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages((Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(new WriteOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages(source, null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(source, new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages(new MemoryStream(source), null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(new MemoryStream(source), new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePages("input.pdf", (Stream)null!, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages("missing.pdf", new ReadOnlyStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(" ", new MemoryStream(), 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages(" ", "out.pdf", 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePages("missing.pdf", " ", 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRange(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRange(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRange("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges((byte[])null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.DuplicatePageRanges(source, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges((Stream)null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(new WriteOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(source, null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(source, new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.DuplicatePageRanges(new MemoryStream(source), null!, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(new MemoryStream(source), new ReadOnlyStream(), PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges(" ", "out.pdf", PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.DuplicatePageRanges("missing.pdf", " ", PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 1, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 2, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, 1, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 5, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePages(source, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages((Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(new WriteOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages(source, null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(source, new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages(new MemoryStream(source), null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(new MemoryStream(source), new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePages("input.pdf", (Stream)null!, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages("missing.pdf", new ReadOnlyStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(" ", new MemoryStream(), 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages(" ", "out.pdf", 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePages("missing.pdf", " ", 1, 2));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(source, 2, 2, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 0, 1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 5, 1, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRange(source, 1, 1, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange((Stream)null!, 1, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(new WriteOnlyStream(), 1, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange(source, null!, 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(source, new ReadOnlyStream(), 4, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRange(new MemoryStream(source), null!, 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(new MemoryStream(source), new ReadOnlyStream(), 4, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange(" ", "out.pdf", 1, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRange("missing.pdf", " ", 1, 1, 2));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, 2, PdfPageRange.From(2, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges((byte[])null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(source, 4, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, 4, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.MovePageRanges(source, 4, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges((Stream)null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(new WriteOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(source, null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(source, new ReadOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.MovePageRanges(new MemoryStream(source), null!, 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(new MemoryStream(source), new ReadOnlyStream(), 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges(" ", "out.pdf", 4, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.MovePageRanges("missing.pdf", " ", 4, PdfPageRange.From(1, 1)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, Array.Empty<int>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPages(source, 1, 2, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages((Stream)null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(new WriteOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages(source, null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(source, new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages(new MemoryStream(source), null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(new MemoryStream(source), new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPages("input.pdf", (Stream)null!, 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages("missing.pdf", new ReadOnlyStream(), 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(" ", new MemoryStream(), 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages(" ", "out.pdf", 1, 2, 3));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPages("missing.pdf", " ", 1, 2, 3));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(source, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(3, 3), new PdfPageRange(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(1, 2), new PdfPageRange(2, 3)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.ReorderPageRanges(source, new PdfPageRange(1, 2), new PdfPageRange(4, 4)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges((Stream)null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(new WriteOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(source, null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(source, new ReadOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.ReorderPageRanges(new MemoryStream(source), null!, new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(new MemoryStream(source), new ReadOnlyStream(), new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges(" ", "out.pdf", new PdfPageRange(1, 3)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.ReorderPageRanges("missing.pdf", " ", new PdfPageRange(1, 3)));

        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(source, 90, 1, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 45, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 90, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePages(source, 90, 4));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages((Stream)null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(new WriteOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages(source, null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(source, new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages(new MemoryStream(source), null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(new MemoryStream(source), new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePages("input.pdf", (Stream)null!, 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages("missing.pdf", new ReadOnlyStream(), 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(" ", new MemoryStream(), 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(" ", "out.pdf", 90, 1));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages("missing.pdf", " ", 90, 1));

        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 3, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 90, 1, 4));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRange(source, 45, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange((Stream)null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(new WriteOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange(source, null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(source, new ReadOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRange(new MemoryStream(source), null!, 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(new MemoryStream(source), new ReadOnlyStream(), 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange(" ", "out.pdf", 90, 1, 2));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRange("missing.pdf", " ", 90, 1, 2));

        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges((byte[])null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(source, 90, (PdfPageRange[])null!));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(source, 90, Array.Empty<PdfPageRange>()));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRanges(source, 90, PdfPageRange.From(1, 4)));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPageEditor.RotatePageRanges(source, 45, PdfPageRange.From(1, 2)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges((Stream)null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(new WriteOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(source, null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(source, new ReadOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentNullException>(() => PdfPageEditor.RotatePageRanges(new MemoryStream(source), null!, 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(new MemoryStream(source), new ReadOnlyStream(), 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges(" ", "out.pdf", 90, PdfPageRange.From(1, 1)));
        Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePageRanges("missing.pdf", " ", 90, PdfPageRange.From(1, 1)));
    }

    [Fact]
    public void PageEditorPathOutputs_RejectDirectoryTargets() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-page-editor-output-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");
        string outputDirectory = Path.Combine(directory, "existing-output");

        try {
            Directory.CreateDirectory(outputDirectory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            var ex = Assert.Throws<ArgumentException>(() => PdfPageEditor.RotatePages(inputPath, outputDirectory, 90));
            Assert.Equal("outputPath", ex.ParamName);
            Assert.Contains("Output path refers to a directory; a file path is required.", ex.Message, StringComparison.Ordinal);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }
}
