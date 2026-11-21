using System.IO.Compression;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderMalformedRelationshipTargetTests {
    private const string Placeholder = "mailto:placeholder@example.test";
    // Not a valid URI on .NET, where System.IO.Packaging cannot read it unless the package is repaired first.
    private const string MalformedTarget = "mailto:one@";

    [Theory]
    [InlineData("docx", ReaderProjection.Full)]
    [InlineData("docx", ReaderProjection.Text)]
    [InlineData("xlsx", ReaderProjection.Full)]
    [InlineData("xlsx", ReaderProjection.Text)]
    [InlineData("pptx", ReaderProjection.Full)]
    [InlineData("pptx", ReaderProjection.Text)]
    public void ReadDocument_ReadsOfficePackagesWithMalformedExternalTargets(string extension, ReaderProjection projection) {
        byte[] package = ReplaceRelationshipTarget(CreateDocument(extension), Placeholder, MalformedTarget);

        OfficeDocumentReadResult result = ReaderTestReaders.All.ReadDocument(package, "contact." + extension,
            new ReaderOptions { Projection = projection });

        Assert.Contains("Contact desk", result.GetText(), StringComparison.Ordinal);
        if (extension == "pptx" && projection == ReaderProjection.Text) {
            // The slide reader opens packages without the repair step, so a target the runtime cannot parse sends
            // the read to the presentation model, which repairs it.
            bool parses = Uri.TryCreate(MalformedTarget, UriKind.RelativeOrAbsolute, out _);
            Assert.Equal(parses, result.CapabilitiesUsed.Contains("officeimo.reader.powerpoint.streamed-text"));
        }
        if (projection == ReaderProjection.Full) {
            Assert.Equal(MalformedTarget, Assert.Single(result.Links).Uri);
        }
    }

    private static byte[] CreateDocument(string extension) {
        using var stream = new MemoryStream();
        switch (extension) {
            case "docx":
                using (WordDocument document = WordDocument.Create(stream)) {
                    document.AddParagraph("Write to ").AddHyperLink("Contact desk", new Uri(Placeholder));
                    document.Save();
                }
                break;
            case "xlsx":
                using (ExcelDocument document = ExcelDocument.Create(stream)) {
                    ExcelSheet sheet = document.AddSheet("Contacts");
                    sheet.Cell(1, 1).SetValue("Contact desk");
                    sheet.SetHyperlink(1, 1, Placeholder, display: "Contact desk");
                    document.Save();
                }
                break;
            default:
                using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                    PowerPointTextBox box = presentation.AddSlide().AddTextBox("Contact desk");
                    box.Paragraphs[0].Runs[0].SetHyperlink(Placeholder);
                    presentation.Save();
                }
                break;
        }
        return stream.ToArray();
    }

    private static byte[] ReplaceRelationshipTarget(byte[] package, string from, string to) {
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        int replaced = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToList()) {
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) xml = reader.ReadToEnd();
                if (!xml.Contains(from, StringComparison.Ordinal)) continue;
                replaced++;
                using Stream output = entry.Open();
                output.SetLength(0);
                byte[] updated = new UTF8Encoding(false).GetBytes(xml.Replace(from, to));
                output.Write(updated, 0, updated.Length);
            }
        }
        Assert.Equal(1, replaced);
        return stream.ToArray();
    }
}
