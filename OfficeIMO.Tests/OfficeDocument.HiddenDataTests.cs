using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentHiddenDataTests {
    [Theory]
    [InlineData("shared.docx", OfficeHiddenDataKind.DocumentProperties)]
    [InlineData("shared.xlsx", OfficeHiddenDataKind.DocumentProperties)]
    [InlineData("shared.pptx", OfficeHiddenDataKind.DocumentProperties)]
    [InlineData("shared.jpg", OfficeHiddenDataKind.ImageMetadata)]
    public void FileRemovalRoutesToTheFormatOwnerAndLeavesTheInputUnchanged(string name, OfficeHiddenDataKind expected) {
        string root = Directory.CreateTempSubdirectory("officeimo-hidden-data-").FullName;
        try {
            string input = Path.Combine(root, name);
            Create(input);
            byte[] original = File.ReadAllBytes(input);

            OfficeHiddenDataReport found = OfficeDocument.Inspect(input, new OfficeInspectOptions { IncludeContent = false, IncludeHiddenData = true }).HiddenData!;
            Assert.Contains(found.Findings, finding => finding.Kind == expected && finding.Outcome == OfficeHiddenDataOutcome.Found);

            string output = Path.Combine(root, "clean" + Path.GetExtension(name));
            OfficeHiddenDataReport removed = OfficeDocument.RemoveHiddenData(input, output);
            Assert.Contains(removed.Findings, finding => finding.Kind == expected && finding.Outcome == OfficeHiddenDataOutcome.Removed);
            Assert.Equal(original, File.ReadAllBytes(input));

            using var stream = File.OpenRead(output);
            OfficeHiddenDataReport after = OfficeDocument.Inspect(stream, output, new OfficeInspectOptions { IncludeContent = false, IncludeHiddenData = true }).HiddenData!;
            Assert.DoesNotContain(after.Findings, finding => finding.Kind == expected);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void Create(string path) {
        switch (Path.GetExtension(path)) {
            case ".docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.Properties.Creator = "Private Author";
                    document.AddParagraph("Shared text");
                    document.Save();
                }
                break;
            case ".xlsx":
                using (ExcelDocument document = ExcelDocument.Create(path)) {
                    document.Properties.Creator = "Private Author";
                    document.AddSheet("Data").Cell(1, 1).SetValue("Shared text");
                    document.Save();
                }
                break;
            case ".pptx":
                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
                    presentation.Properties.Creator = "Private Author";
                    presentation.AddSlide();
                    presentation.Save();
                }
                break;
            default:
                byte[] jpeg = OfficeJpegCodec.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
                OfficeImageMetadata metadata = OfficeImageMetadata.Read(jpeg);
                metadata.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
                File.WriteAllBytes(path, OfficeImageMetadata.Apply(jpeg, metadata));
                break;
        }
    }
}
