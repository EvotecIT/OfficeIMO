using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.OneNote;

var section = new OneNoteSection { Name = "Packed section" };
var page = new OneNotePage { Title = "Packed page" };
var paragraph = new OneNoteParagraph();
var bold = new OneNoteTextRun { Text = "Packed OneNote content" };
bold.Style.Bold = true;
paragraph.Runs.Add(bold);
paragraph.Tags.Add(new OneNoteTag { Label = "To Do", IsCheckable = true });
page.DirectContent.Add(paragraph);
page.DirectContent.Add(new OneNoteEmbeddedFile {
    FileName = "proof.txt",
    MediaType = "text/plain",
    Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3, 4 })
});
section.Pages.Add(page);

byte[] sectionBytes = OneNoteSectionWriter.Write(section);
OneNoteSection reopened = OneNoteSectionReader.Read(new MemoryStream(sectionBytes));
reopened.Pages[0].Title = "Edited packed page";
OneNoteSection edited = OneNoteSectionReader.Read(new MemoryStream(OneNoteSectionWriter.Write(reopened)));
if (!string.Equals(edited.Pages[0].Title, "Edited packed page", StringComparison.Ordinal)) {
    throw new InvalidOperationException("The packed core package failed native write/edit/read validation.");
}

string markdown = edited.ToMarkdown();
string html = edited.ToHtmlDocument();
byte[] pdf = edited.ToPdf();
if (markdown.IndexOf("Packed OneNote content", StringComparison.Ordinal) < 0 ||
    html.IndexOf("Packed OneNote content", StringComparison.Ordinal) < 0 ||
    pdf.Length < 5 ||
    pdf[0] != (byte)'%' ||
    pdf[1] != (byte)'P' ||
    pdf[2] != (byte)'D' ||
    pdf[3] != (byte)'F') {
    throw new InvalidOperationException("The packed Markdown, HTML, or PDF conversion graph failed.");
}

var notebook = new OneNoteNotebook { Name = "Packed notebook" };
notebook.Sections.Add(edited);
byte[] packageBytes = OneNotePackageWriter.Write(notebook);
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
using var packageStream = new MemoryStream(packageBytes);
OfficeDocumentReadResult readerResult = await reader.ReadDocumentAsync(packageStream, "packed.onepkg");
if (readerResult.Kind != ReaderInputKind.OneNote ||
    !readerResult.Chunks.Any(chunk => chunk.Text.IndexOf("Packed OneNote content", StringComparison.Ordinal) >= 0) ||
    !readerResult.Assets.Any(asset => string.Equals(asset.FileName, "proof.txt", StringComparison.Ordinal))) {
    throw new InvalidOperationException("The packed Reader adapter failed structured .onepkg ingestion.");
}

Console.WriteLine($"OfficeIMO.OneNote package-family smoke passed on {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}.");
