using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>Preservation contracts at the native boundary: omitted values and independently identified outputs.</summary>
public sealed class StudioNativePreservationTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-preservation-" + Guid.NewGuid().ToString("N"));

    public StudioNativePreservationTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, true);

    private string CreatePdf() {
        string path = Path.Combine(_root, "source.pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(300, 400))).Save(path);
        return path;
    }

    [Fact]
    public async Task Cropping_uses_the_dimensions_of_the_rotated_page() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(), null, default);
        await document.RotatePagesAsync([1], 90, default);
        await document.EditAsync(new StudioNativeEditRequest(Crop: new StudioNativeCropEdit([1], Left: 160, Top: 0, Right: 160, Bottom: 0)), default);
        StudioNativePageInfo page = Assert.Single((await document.GetInfoAsync(default)).Pages);
        Assert.Equal(90, page.Rotation);
        Assert.Equal(300D, page.Width, 4);
        Assert.Equal(80D, page.Height, 4);
    }

    [Fact]
    public async Task A_text_only_annotation_edit_preserves_absent_color_and_author() {
        string path = CreatePdf();
        byte[] annotated = PdfDocument.Load(File.ReadAllBytes(path)).Annotations.Add(new PdfAnnotationCreateOptions {
            Contents = "Original", Rectangle = [40, 50, 64, 74], IconName = "Comment"
        }).Bytes;
        File.WriteAllBytes(path, annotated);
        PdfAnnotation before = Assert.Single(PdfDocument.Load(annotated).Inspect().Annotations);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        await document.UpdateAnnotationAsync(new StudioNativeAnnotationUpdate(before.ObjectNumber!.Value, Contents: "Revised"), default);
        string saved = Path.Combine(_root, "annotated.pdf");
        await document.SaveCopyAsync(saved, default);
        PdfAnnotation after = Assert.Single(PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations);
        Assert.Equal("Revised", after.Contents);
        Assert.Equal(before.Title, after.Title);
        Assert.Equal(before.Color, after.Color);
    }

    [Fact]
    public async Task A_non_replacing_extraction_preserves_an_existing_destination() {
        using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(CreatePdf(), default);
        string output = Path.Combine(_root, "taken.pdf");
        byte[] original = [1, 2, 3, 4];
        await File.WriteAllBytesAsync(output, original);
        await Assert.ThrowsAsync<IOException>(() => workspace.ExtractToLocalFileAsync([1], output, default, replaceExisting: false));
        Assert.Equal(original, await File.ReadAllBytesAsync(output));
    }

    [Fact]
    public async Task Duplicate_attachment_filenames_remain_individually_addressable() {
        string path = Path.Combine(_root, "attachments.pdf");
        File.WriteAllBytes(path, AttachmentFixture());
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        StudioNativeAttachment[] attachments = (await document.GetStructureAsync(default)).Attachments.ToArray();
        Assert.Equal(2, attachments.Length);
        Assert.All(attachments, attachment => Assert.Equal("same.txt", attachment.FileName));
        Assert.NotEqual(attachments[0].Id, attachments[1].Id);
        string output = Path.Combine(_root, "selected.txt");
        await Assert.ThrowsAsync<ArgumentException>(() => document.WriteAsync(
            new StudioNativeWriteRequest(Attachment: new StudioNativeAttachmentSave("same.txt", output)), default));
        await document.WriteAsync(new StudioNativeWriteRequest(Attachment:
            new StudioNativeAttachmentSave("same.txt", output, attachments[1].Id)), default);
        Assert.Equal("second", File.ReadAllText(output));
        await document.EditAsync(new StudioNativeEditRequest(RemoveAttachment: attachments[1].Id), default);
        Assert.Equal("same.txt", Assert.Single((await document.GetStructureAsync(default)).Attachments).FileName);
        await document.WriteAsync(new StudioNativeWriteRequest(Attachment: new StudioNativeAttachmentSave("same.txt", output)), default);
        Assert.Equal("first", File.ReadAllText(output));
    }

    private static byte[] AttachmentFixture() {
        string[] objects = [
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(key-a) 5 0 R (key-b) 7 0 R] >> >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\n\nendstream",
            "<< /Type /Filespec /F (same.txt) /UF (same.txt) /EF << /F 6 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 5 >>\nstream\nfirst\nendstream",
            "<< /Type /Filespec /F (same.txt) /UF (same.txt) /EF << /F 8 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 6 >>\nstream\nsecond\nendstream"
        ];
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        foreach (string value in objects) {
            offsets.Add(pdf.Length);
            pdf.Append(offsets.Count).Append(" 0 obj\n").Append(value).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer\n<< /Size ").Append(objects.Length + 1).Append(" /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
