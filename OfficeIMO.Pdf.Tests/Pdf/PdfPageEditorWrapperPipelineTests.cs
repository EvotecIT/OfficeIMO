using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageEditorTests {

    [Fact]
    public void PageEditorPathInputs_ReturnBytesForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-editor-path-bytes-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            byte[] deleted = PdfPageEditor.DeletePages(inputPath, 2);
            string deletedText = NormalizeExtractedText(PdfReadDocument.Open(deleted).ExtractText());
            Assert.Contains("Firstpagemarker", deletedText);
            Assert.DoesNotContain("Secondpagemarker", deletedText);
            Assert.Contains("Thirdpagemarker", deletedText);

            byte[] deletedRange = PdfPageEditor.DeletePageRange(inputPath, 1, 2);
            var deletedRangeRead = PdfReadDocument.Open(deletedRange);
            Assert.Single(deletedRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedRangeRead.Pages[0].ExtractText()));

            byte[] deletedModelRange = PdfPageEditor.DeletePageRange(inputPath, PdfPageRange.From(1, 2));
            var deletedModelRangeRead = PdfReadDocument.Open(deletedModelRange);
            Assert.Single(deletedModelRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedModelRangeRead.Pages[0].ExtractText()));

            byte[] deletedRanges = PdfPageEditor.DeletePageRanges(inputPath, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));
            var deletedRangesRead = PdfReadDocument.Open(deletedRanges);
            Assert.Single(deletedRangesRead.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(deletedRangesRead.Pages[0].ExtractText()));

            byte[] duplicated = PdfPageEditor.DuplicatePages(inputPath, 3);
            var duplicatedRead = PdfReadDocument.Open(duplicated);
            Assert.Equal(4, duplicatedRead.Pages.Count);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[3].ExtractText()));

            byte[] duplicatedRange = PdfPageEditor.DuplicatePageRange(inputPath, 1, 2);
            var duplicatedRangeRead = PdfReadDocument.Open(duplicatedRange);
            Assert.Equal(5, duplicatedRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[3].ExtractText()));

            byte[] duplicatedModelRange = PdfPageEditor.DuplicatePageRange(inputPath, PdfPageRange.From(1, 2));
            var duplicatedModelRangeRead = PdfReadDocument.Open(duplicatedModelRange);
            Assert.Equal(5, duplicatedModelRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[3].ExtractText()));

            byte[] duplicatedRanges = PdfPageEditor.DuplicatePageRanges(inputPath, PdfPageRange.ParseMany("1,3"));
            var duplicatedRangesRead = PdfReadDocument.Open(duplicatedRanges);
            Assert.Equal(5, duplicatedRangesRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[3].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[4].ExtractText()));

            byte[] moved = PdfPageEditor.MovePages(inputPath, 1, 2);
            var movedRead = PdfReadDocument.Open(moved);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRead.Pages[1].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRead.Pages[2].ExtractText()));

            byte[] movedRange = PdfPageEditor.MovePageRange(inputPath, 4, 1, 2);
            var movedRangeRead = PdfReadDocument.Open(movedRange);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangeRead.Pages[2].ExtractText()));

            byte[] movedModelRange = PdfPageEditor.MovePageRange(inputPath, 4, PdfPageRange.From(1, 2));
            var movedModelRangeRead = PdfReadDocument.Open(movedModelRange);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[2].ExtractText()));

            byte[] movedRanges = PdfPageEditor.MovePageRanges(inputPath, 4, PdfPageRange.ParseMany("1-2,2"));
            var movedRangesRead = PdfReadDocument.Open(movedRanges);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangesRead.Pages[2].ExtractText()));

            byte[] reordered = PdfPageEditor.ReorderPages(inputPath, 3, 1, 2);
            string reorderedText = NormalizeExtractedText(PdfReadDocument.Open(reordered).ExtractText());
            Assert.True(reorderedText.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal));
            Assert.True(reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Secondpagemarker", StringComparison.Ordinal));

            byte[] rotated = PdfPageEditor.RotatePages(inputPath, 90, 1);
            PdfDocumentInfo rotatedInfo = PdfInspector.Inspect(rotated);
            Assert.Equal(90, rotatedInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[1].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[2].RotationDegrees);

            byte[] rotatedRange = PdfPageEditor.RotatePageRange(inputPath, 180, 2, 3);
            PdfDocumentInfo rotatedRangeInfo = PdfInspector.Inspect(rotatedRange);
            Assert.Equal(0, rotatedRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[2].RotationDegrees);

            byte[] rotatedModelRange = PdfPageEditor.RotatePageRange(inputPath, 180, PdfPageRange.From(2, 3));
            PdfDocumentInfo rotatedModelRangeInfo = PdfInspector.Inspect(rotatedModelRange);
            Assert.Equal(0, rotatedModelRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[2].RotationDegrees);

            byte[] rotatedRanges = PdfPageEditor.RotatePageRanges(inputPath, 270, PdfPageRange.ParseMany("1,3"));
            PdfDocumentInfo rotatedRangesInfo = PdfInspector.Inspect(rotatedRanges);
            Assert.Equal(270, rotatedRangesInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedRangesInfo.Pages[1].RotationDegrees);
            Assert.Equal(270, rotatedRangesInfo.Pages[2].RotationDegrees);

            byte[] resized = PdfPageEditor.ResizePages(inputPath, PageSizes.Letter, 3);
            PdfDocumentInfo resizedInfo = PdfInspector.Inspect(resized);
            Assert.Equal(3, resizedInfo.PageCount);
            Assert.Equal(595, Math.Round(resizedInfo.Pages[0].Width));
            Assert.Equal(842, Math.Round(resizedInfo.Pages[0].Height));
            Assert.Equal(612, Math.Round(resizedInfo.Pages[2].Width));
            Assert.Equal(792, Math.Round(resizedInfo.Pages[2].Height));
            Assert.Contains("Thirdpagemarker", NormalizeExtractedText(PdfReadDocument.Open(resized).ExtractText()));
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void PageEditorPathInputs_WriteToOutputStreamsForWrapperPipelines() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-pdf-editor-path-stream-" + Guid.NewGuid().ToString("N"));
        string inputPath = Path.Combine(directory, "input.pdf");

        try {
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(inputPath, BuildThreePagePdf());

            using var deletedOutput = CreateOutputStream(out int deletedPrefixLength);
            PdfPageEditor.DeletePages(inputPath, deletedOutput, 2);
            string deletedText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(deletedOutput, deletedPrefixLength)).ExtractText());
            Assert.Contains("Firstpagemarker", deletedText);
            Assert.DoesNotContain("Secondpagemarker", deletedText);
            Assert.Contains("Thirdpagemarker", deletedText);

            using var deletedRangeOutput = CreateOutputStream(out int deletedRangePrefixLength);
            PdfPageEditor.DeletePageRange(inputPath, deletedRangeOutput, 1, 2);
            var deletedRangeRead = PdfReadDocument.Open(GetOutputPayload(deletedRangeOutput, deletedRangePrefixLength));
            Assert.Single(deletedRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedRangeRead.Pages[0].ExtractText()));

            using var deletedModelRangeOutput = CreateOutputStream(out int deletedModelRangePrefixLength);
            PdfPageEditor.DeletePageRange(inputPath, deletedModelRangeOutput, PdfPageRange.From(1, 2));
            var deletedModelRangeRead = PdfReadDocument.Open(GetOutputPayload(deletedModelRangeOutput, deletedModelRangePrefixLength));
            Assert.Single(deletedModelRangeRead.Pages);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(deletedModelRangeRead.Pages[0].ExtractText()));

            using var deletedRangesOutput = CreateOutputStream(out int deletedRangesPrefixLength);
            PdfPageEditor.DeletePageRanges(inputPath, deletedRangesOutput, PdfPageRange.From(1, 1), PdfPageRange.From(3, 3));
            var deletedRangesRead = PdfReadDocument.Open(GetOutputPayload(deletedRangesOutput, deletedRangesPrefixLength));
            Assert.Single(deletedRangesRead.Pages);
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(deletedRangesRead.Pages[0].ExtractText()));

            using var duplicatedOutput = CreateOutputStream(out int duplicatedPrefixLength);
            PdfPageEditor.DuplicatePages(inputPath, duplicatedOutput, 3);
            var duplicatedRead = PdfReadDocument.Open(GetOutputPayload(duplicatedOutput, duplicatedPrefixLength));
            Assert.Equal(4, duplicatedRead.Pages.Count);
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRead.Pages[3].ExtractText()));

            using var duplicatedRangeOutput = CreateOutputStream(out int duplicatedRangePrefixLength);
            PdfPageEditor.DuplicatePageRange(inputPath, duplicatedRangeOutput, 1, 2);
            var duplicatedRangeRead = PdfReadDocument.Open(GetOutputPayload(duplicatedRangeOutput, duplicatedRangePrefixLength));
            Assert.Equal(5, duplicatedRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangeRead.Pages[3].ExtractText()));

            using var duplicatedModelRangeOutput = CreateOutputStream(out int duplicatedModelRangePrefixLength);
            PdfPageEditor.DuplicatePageRange(inputPath, duplicatedModelRangeOutput, PdfPageRange.From(1, 2));
            var duplicatedModelRangeRead = PdfReadDocument.Open(GetOutputPayload(duplicatedModelRangeOutput, duplicatedModelRangePrefixLength));
            Assert.Equal(5, duplicatedModelRangeRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[2].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedModelRangeRead.Pages[3].ExtractText()));

            using var duplicatedRangesOutput = CreateOutputStream(out int duplicatedRangesPrefixLength);
            PdfPageEditor.DuplicatePageRanges(inputPath, duplicatedRangesOutput, PdfPageRange.ParseMany("1,3"));
            var duplicatedRangesRead = PdfReadDocument.Open(GetOutputPayload(duplicatedRangesOutput, duplicatedRangesPrefixLength));
            Assert.Equal(5, duplicatedRangesRead.Pages.Count);
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[2].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[3].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(duplicatedRangesRead.Pages[4].ExtractText()));

            using var movedOutput = CreateOutputStream(out int movedPrefixLength);
            PdfPageEditor.MovePages(inputPath, movedOutput, 1, 2);
            var movedRead = PdfReadDocument.Open(GetOutputPayload(movedOutput, movedPrefixLength));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRead.Pages[1].ExtractText()));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRead.Pages[2].ExtractText()));

            using var movedRangeOutput = CreateOutputStream(out int movedRangePrefixLength);
            PdfPageEditor.MovePageRange(inputPath, movedRangeOutput, 4, 1, 2);
            var movedRangeRead = PdfReadDocument.Open(GetOutputPayload(movedRangeOutput, movedRangePrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangeRead.Pages[2].ExtractText()));

            using var movedModelRangeOutput = CreateOutputStream(out int movedModelRangePrefixLength);
            PdfPageEditor.MovePageRange(inputPath, movedModelRangeOutput, 4, PdfPageRange.From(1, 2));
            var movedModelRangeRead = PdfReadDocument.Open(GetOutputPayload(movedModelRangeOutput, movedModelRangePrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedModelRangeRead.Pages[2].ExtractText()));

            using var movedRangesOutput = CreateOutputStream(out int movedRangesPrefixLength);
            PdfPageEditor.MovePageRanges(inputPath, movedRangesOutput, 4, PdfPageRange.ParseMany("1-2,2"));
            var movedRangesRead = PdfReadDocument.Open(GetOutputPayload(movedRangesOutput, movedRangesPrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(movedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(movedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(movedRangesRead.Pages[2].ExtractText()));

            using var reorderedOutput = CreateOutputStream(out int reorderedPrefixLength);
            PdfPageEditor.ReorderPages(inputPath, reorderedOutput, 3, 1, 2);
            string reorderedText = NormalizeExtractedText(PdfReadDocument.Open(GetOutputPayload(reorderedOutput, reorderedPrefixLength)).ExtractText());
            Assert.True(reorderedText.IndexOf("Thirdpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal));
            Assert.True(reorderedText.IndexOf("Firstpagemarker", StringComparison.Ordinal) < reorderedText.IndexOf("Secondpagemarker", StringComparison.Ordinal));

            using var reorderedRangesOutput = CreateOutputStream(out int reorderedRangesPrefixLength);
            PdfPageEditor.ReorderPageRanges(inputPath, reorderedRangesOutput, PdfPageRange.From(3, 3), PdfPageRange.From(1, 2));
            var reorderedRangesRead = PdfReadDocument.Open(GetOutputPayload(reorderedRangesOutput, reorderedRangesPrefixLength));
            Assert.Equal("Thirdpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[0].ExtractText()));
            Assert.Equal("Firstpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[1].ExtractText()));
            Assert.Equal("Secondpagemarker", NormalizeExtractedText(reorderedRangesRead.Pages[2].ExtractText()));

            using var rotatedOutput = CreateOutputStream(out int rotatedPrefixLength);
            PdfPageEditor.RotatePages(inputPath, rotatedOutput, 90, 1);
            PdfDocumentInfo rotatedInfo = PdfInspector.Inspect(GetOutputPayload(rotatedOutput, rotatedPrefixLength));
            Assert.Equal(90, rotatedInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[1].RotationDegrees);
            Assert.Equal(0, rotatedInfo.Pages[2].RotationDegrees);

            using var rotatedRangeOutput = CreateOutputStream(out int rotatedRangePrefixLength);
            PdfPageEditor.RotatePageRange(inputPath, rotatedRangeOutput, 180, 2, 3);
            PdfDocumentInfo rotatedRangeInfo = PdfInspector.Inspect(GetOutputPayload(rotatedRangeOutput, rotatedRangePrefixLength));
            Assert.Equal(0, rotatedRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedRangeInfo.Pages[2].RotationDegrees);

            using var rotatedModelRangeOutput = CreateOutputStream(out int rotatedModelRangePrefixLength);
            PdfPageEditor.RotatePageRange(inputPath, rotatedModelRangeOutput, 180, PdfPageRange.From(2, 3));
            PdfDocumentInfo rotatedModelRangeInfo = PdfInspector.Inspect(GetOutputPayload(rotatedModelRangeOutput, rotatedModelRangePrefixLength));
            Assert.Equal(0, rotatedModelRangeInfo.Pages[0].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[1].RotationDegrees);
            Assert.Equal(180, rotatedModelRangeInfo.Pages[2].RotationDegrees);

            using var rotatedRangesOutput = CreateOutputStream(out int rotatedRangesPrefixLength);
            PdfPageEditor.RotatePageRanges(inputPath, rotatedRangesOutput, 270, PdfPageRange.ParseMany("1,3"));
            PdfDocumentInfo rotatedRangesInfo = PdfInspector.Inspect(GetOutputPayload(rotatedRangesOutput, rotatedRangesPrefixLength));
            Assert.Equal(270, rotatedRangesInfo.Pages[0].RotationDegrees);
            Assert.Equal(0, rotatedRangesInfo.Pages[1].RotationDegrees);
            Assert.Equal(270, rotatedRangesInfo.Pages[2].RotationDegrees);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

}
