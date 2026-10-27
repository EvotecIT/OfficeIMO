using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.OpenDocument;
using System.Text;
using OfficeIMO.Word;
using OfficeIMO.ContentSafety;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderInspectionResourceLimitEvidenceTests {
    [Fact]
    public void NativeSafetyFindingLimitCannotBecomeASkippedZipMember() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().AddZipHandler().Build();
        using var zip = new MemoryStream();
        using (var archive = new ZipArchive(zip, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream output = archive.CreateEntry("hidden.docx").Open(); byte[] source = HiddenWord(); output.Write(source, 0, source.Length);
        }
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(zip.ToArray(), "hidden.zip", Bounded()));
        Assert.Equal("Inspection.MaxRecords", error.LimitName);
    }

    [Theory]
    [InlineData("detailed")]
    [InlineData("ordered")]
    [InlineData("completed")]
    [InlineData("folder")]
    public async Task NativeSafetyFindingLimitRemainsTerminalAcrossAggregateReadRoutes(string route) {
        string folder = Path.Combine(Path.GetTempPath(), "reader-native-safety-limit-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "hidden.docx"); File.WriteAllBytes(path, HiddenWord());
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
            ReaderResourceLimitException error = await Assert.ThrowsAsync<ReaderResourceLimitException>(async () => {
                if (route == "detailed") await reader.ReadDocumentsDetailedAsync(new[] { path }, Bounded());
                else if (route == "ordered") await reader.ReadDocumentsAsync(new[] { path }, Bounded());
                else if (route == "completed") await reader.ReadDocumentsAsCompletedAsync(new[] { path }, _ => { }, Bounded());
                else reader.ReadFolderDetailed(folder, options: Bounded());
            });
            Assert.Equal("Inspection.MaxRecords", error.LimitName);
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task MalformedDocumentRemainsAnOrdinaryFailedBatchOutcome() {
        string folder = Path.Combine(Path.GetTempPath(), "reader-native-malformed-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "malformed.docx"); File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddWordHandler().Build();
            ReaderDocumentReadOutcome outcome = Assert.Single(await reader.ReadDocumentsDetailedAsync(new[] { path }, Bounded()));
            Assert.NotNull(outcome.Error); Assert.IsNotType<ReaderResourceLimitException>(outcome.Error);
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void NativeUnicodeAndOtherAdmissionCeilingsRetainTypedLimitIdentity() {
        var builder = new OfficeContentSafetyBuilder("TEST", new OfficeContentSafetyOptions { MaxFindings = 1 });
        OfficeInspectionResourceLimitException unicode = Assert.Throws<OfficeInspectionResourceLimitException>(() => builder.InspectVisibleText("Body", "\u200B\u200B"));
        Assert.Equal("content-safety", unicode.Owner); Assert.Equal(nameof(OfficeContentSafetyOptions.MaxFindings), unicode.LimitName);
        Assert.Equal(1, unicode.Maximum); Assert.IsType<OfficeInspectionResourceLimitException>(unicode.InnerException);
        OfficeInspectionResourceLimitException characters = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            OfficeContentSafetyInputGuard.ValidateText("source", new OfficeContentSafetyOptions { MaxCharacters = 1 }));
        Assert.Equal(nameof(OfficeContentSafetyOptions.MaxCharacters), characters.LimitName);
        OfficeInspectionResourceLimitException bytes = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            OfficeContentSafetyInputGuard.ValidateBytes(new byte[] { 1, 2 }, new OfficeContentSafetyOptions { MaxInputBytes = 1 }));
        Assert.Equal(nameof(OfficeContentSafetyOptions.MaxInputBytes), bytes.LimitName);
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    [InlineData("odp")]
    [InlineData("rtf")]
    public void EveryNativeSafetyProjectionPreservesItsTerminalOwnerLimit(string extension) {
        byte[] bytes;
        if (extension == "xlsx") {
            using var memory = new MemoryStream();
            using (ExcelDocument document = ExcelDocument.Create(memory)) {
                ExcelSheet sheet = document.AddWorksheet("Hidden source");
                for (int index = 1; index <= 4; index++) sheet.Cell(index, 1, "Hidden source " + index);
                foreach (var row in document.OpenXmlDocument!.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>()) row.Hidden = true;
                document.Save();
            }
            bytes = memory.ToArray();
        } else if (extension == "pptx") {
            using var memory = new MemoryStream(); using PowerPointDocument document = PowerPointDocument.Create(memory);
            PowerPointSlide slide = document.AddSlide(); for (int index = 1; index <= 4; index++) slide.AddTextBox("Hidden source " + index).Hidden = true;
            bytes = document.ToBytes();
        } else if (extension == "odp") {
            OdpPresentation document = OdpPresentation.Create(); OdpSlide slide = document.AddSlide();
            for (int index = 1; index <= 4; index++) slide.AddTextBox(OdfRect.FromCentimeters(1, index, 2, 1), "Hidden source " + index).Hidden = true;
            bytes = document.ToBytes();
        } else bytes = Encoding.ASCII.GetBytes(@"{\rtf1\ansi {\v First}\par {\v Second}\par {\v Third}\par {\v Fourth}}");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddExcelHandler().AddPowerPointHandler().AddOpenDocumentHandler().AddRtfHandler().Build();
        ReaderResourceLimitException error = Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(bytes, "hidden." + extension, Bounded()));
        Assert.Equal("Inspection.MaxRecords", error.LimitName);
        Assert.IsType<OfficeInspectionResourceLimitException>(error.InnerException);
    }

    private static ReaderOptions Bounded() => new() { Inspection = new ReaderInspectionOptions {
        IncludeComments = false, IncludeRevisions = false, IncludeVisibility = false, IncludeFormulas = false,
        IncludePackageContent = false, IncludeContentSafety = true, MaxRecords = 3
    }};

    private static byte[] HiddenWord() {
        using var memory = new MemoryStream();
        using (WordDocument document = WordDocument.Create(memory)) {
            document.OpenXmlDocument!.MainDocumentPart!.Document!.Body!.Append(new W.Paragraph(
                Enumerable.Range(1, 4).Select(index => new W.Run(new W.RunProperties(new W.Vanish()), new W.Text("Hidden source " + index)))));
            document.Save();
        }
        return memory.ToArray();
    }
}
