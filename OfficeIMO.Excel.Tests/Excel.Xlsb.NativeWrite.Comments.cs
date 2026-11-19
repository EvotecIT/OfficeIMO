using OfficeIMO.Excel;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Legacy cell comments (notes) written, read and rewritten by the native XLSB writers.</summary>
public sealed class ExcelXlsbCommentTests {
    private const string CommentsPart = "xl/comments1.bin";
    private const string VmlPart = "xl/drawings/vmlDrawing1.vml";

    [Fact]
    public void Comments_from_several_authors_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Value");
            sheet.SetComment("B2", "Check this total", "Alice");
            sheet.SetComment("D5", "Line one\nLine two", "Bob");
            sheet.SetComment("A10", "Second note", "Alice");
            document.AddSheet("Empty").Cell(1, 1).SetValue(1);
        });

        byte[] xlsb = SaveAsXlsb(xlsx);

        Assert.True(HasEntry(xlsb, CommentsPart));
        Assert.True(HasEntry(xlsb, VmlPart));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        Assert.Equal(
            new[] { "B2|Alice|Check this total", "D5|Bob|Line one\nLine two", "A10|Alice|Second note" },
            reloaded["Data"].GetComments().Select(comment => comment.CellReference + "|" + comment.Author + "|" + comment.Text));
        Assert.Empty(reloaded["Empty"].GetComments());
    }

    [Fact]
    public void Formatted_comment_runs_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.SetCommentRichText("C3", new[] {
                new ExcelRichTextRun("Warning: ") { Bold = true },
                new ExcelRichTextRun("check "),
                new ExcelRichTextRun("inputs") { Italic = true, FontFamily = "Tahoma" },
            }, "Reviewer");
        });

        byte[] xlsb = SaveAsXlsb(xlsx);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        ExcelCommentInfo comment = Assert.Single(reloaded["Data"].GetComments());
        Assert.Equal("Warning: check inputs", comment.Text);
        Assert.Equal(new[] { "Warning: ", "check ", "inputs" }, comment.RichTextRuns.Select(run => run.Text));
        Assert.True(comment.RichTextRuns[0].Bold);
        Assert.False(comment.RichTextRuns[1].Bold);
        Assert.False(comment.RichTextRuns[1].Italic);
        Assert.True(comment.RichTextRuns[2].Italic);
        Assert.Equal("Tahoma", comment.RichTextRuns[2].FontFamily);
    }

    [Fact]
    public void Saving_a_loaded_xlsb_with_a_cell_edit_copies_its_comment_parts() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Before");
            sheet.SetComment("B2", "Keep me", "Alice");
        }));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].Cell(1, 1).SetValue("After");
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.Equal(ReadEntry(source, CommentsPart), ReadEntry(edited, CommentsPart));
        Assert.Equal(ReadEntry(source, VmlPart), ReadEntry(edited, VmlPart));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Assert.True(reloaded["Data"].TryGetCellText(1, 1, out string? a1));
        Assert.Equal("After", a1);
        Assert.Equal("Keep me", Assert.Single(reloaded["Data"].GetComments()).Text);
    }

    [Fact]
    public void Editing_comments_of_a_loaded_xlsb_rewrites_its_comment_parts() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.SetComment("B2", "Original", "Alice");
            sheet.SetComment("C3", "Unchanged", "Bob");
        }));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].SetComment("B2", "Revised", "Carol");
            loaded["Data"].SetComment("E7", "Added", "Bob");
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        Assert.NotEqual(ReadEntry(source, CommentsPart), ReadEntry(edited, CommentsPart));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Assert.Equal(
            new[] { "C3|Bob|Unchanged", "B2|Carol|Revised", "E7|Bob|Added" },
            reloaded["Data"].GetComments().Select(comment => comment.CellReference + "|" + comment.Author + "|" + comment.Text));
    }

    [Fact]
    public void Adding_comments_to_a_loaded_xlsb_sheet_without_comments_is_refused() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue(1)));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        loaded["Data"].SetComment("A1", "New", "Alice");

        NotSupportedException error = Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("had no comments", error.Message);
    }

    [Fact]
    public void Threaded_comments_are_refused_by_native_xlsb_generation() {
        byte[] xlsx = CreateXlsx(document => document.AddSheet("Data").AddThreadedComment("A1", "Discuss", "Alice"));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        NotSupportedException error = Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("threaded comments", error.Message);
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static bool HasEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.GetEntry(name) != null;
    }

    private static byte[] ReadEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        using (Stream entry = archive.GetEntry(name)!.Open()) entry.CopyTo(output);
        return output.ToArray();
    }
}
