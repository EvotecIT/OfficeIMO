using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceBufferSubtypeTests {
    [Fact]
    public void RenamedBinaryWorkbookIsRejectedBeforeInspectionOrRemoval() {
        using var workbook = ExcelDocument.Create();
        workbook.AddWorksheet("Data").Cell(1, 1, "Binary workbook");
        byte[] bytes = workbook.ToBytes(ExcelFileFormat.Xlsb);
        Verify(bytes, "renamed.xlsx", workflowSupported: false, ownerSupported: true,
            (data, options) => ExcelDocument.InspectProvenance(data, "owner.xlsb", options));
    }

    [Fact]
    public void QualifiedOfficeExtensionsRejectRenamedSubtypesWithoutChangingOwnerSupport() {
        foreach (var type in Enum.GetValues<WordprocessingDocumentType>()) {
            using var stream = new MemoryStream();
            using (var package = WordprocessingDocument.Create(stream, type))
                package.AddMainDocumentPart().Document = new DocumentFormat.OpenXml.Wordprocessing.Document(new DocumentFormat.OpenXml.Wordprocessing.Body());
            bool supported = type == WordprocessingDocumentType.Document;
            Verify(stream.ToArray(), "renamed.docx", supported, supported,
                (bytes, options) => WordDocument.InspectProvenance(bytes, "owner.docx", options));
        }
        foreach (var type in Enum.GetValues<SpreadsheetDocumentType>()) {
            using var stream = new MemoryStream();
            using (var package = SpreadsheetDocument.Create(stream, type))
                package.AddWorkbookPart().Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            bool supported = type == SpreadsheetDocumentType.Workbook;
            Verify(stream.ToArray(), "renamed.xlsx", supported, supported,
                (bytes, options) => ExcelDocument.InspectProvenance(bytes, "owner.xlsx", options));
        }
        foreach (var type in Enum.GetValues<PresentationDocumentType>()) {
            using var stream = new MemoryStream();
            using (var package = PresentationDocument.Create(stream, type))
                package.AddPresentationPart().Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();
            bool supported = type == PresentationDocumentType.Presentation;
            Verify(stream.ToArray(), "renamed.pptx", supported, supported,
                (bytes, options) => PowerPointPresentation.InspectProvenance(bytes, "owner.pptx", options));
        }
    }

    private static void Verify(
        byte[] bytes,
        string name,
        bool workflowSupported,
        bool ownerSupported,
        Func<byte[], OfficeProvenanceOptions, OfficeProvenanceReport> owner) {
        var inspection = new OfficeProvenanceOptions();
        var removal = new OfficeProvenanceRemovalOptions();
        if (workflowSupported) {
            Assert.Empty(OfficeProvenanceBufferWorkflow.Inspect(bytes, name, inspection).Evidence);
            Assert.Equal(bytes, OfficeProvenanceBufferWorkflow.Remove(bytes, name, removal).ToArray());
        } else {
            Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Inspect(bytes, name, inspection));
            Assert.Throws<InvalidDataException>(() => OfficeProvenanceBufferWorkflow.Remove(bytes, name, removal));
        }
        if (ownerSupported) {
            Assert.Empty(owner(bytes, inspection).Evidence);
            Assert.Empty(owner(bytes, removal.Limits).Evidence);
        } else {
            Assert.Throws<InvalidDataException>(() => owner(bytes, inspection));
            Assert.Throws<InvalidDataException>(() => owner(bytes, removal.Limits));
        }
    }
}
