using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.Pdf;
using OpenMcdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class LegacyDocPdfStageTests {
    [Fact]
    public void UnsupportedLegacyStorageBlocksByDefaultAndRemainsVisibleWhenAccepted() {
        byte[] source = CreateLegacySource();
        var settings = new LegacyDocImportOptions { ReportUnsupportedContent = false };
        Assert.Throws<InvalidOperationException>(() => LegacyDocPdfConverter.ToPdfDocumentResult(new MemoryStream(source), importOptions: settings));
        var accepted = LegacyDocPdfConverter.ToPdfDocumentResult(new MemoryStream(source), importOptions: settings,
            lossPolicy: OfficeConversionLossPolicy.Allow);
        Assert.True(Assert.Single(accepted.SourceConversionReports).HasLoss);
        using var destination = new MemoryStream();
        PdfSaveResult output = accepted.SaveResult(destination);
        output.RequireSuccess();
        Assert.True(output.HasLoss);
        Assert.Equal(2, output.ConversionReports.Count);
        Assert.Contains(output.Report.Warnings, warning => warning.Code == "NativeFontFamilySubstituted");
        Assert.Contains(accepted.FidelityDiagnostics, finding => finding.LossKind != OfficeConversionLossKind.None);
    }

    internal static byte[] CreateLegacySource() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-legacy-pdf-" + Guid.NewGuid().ToString("N") + ".doc");
        byte[] source;
        try {
            using var word = WordDocument.Create(); word.AddParagraph("Preserved DOC text"); word.Save(path);
            using var package = new MemoryStream();
            package.Write(File.ReadAllBytes(path)); package.Position = 0;
            using (RootStorage root = RootStorage.Open(package, StorageModeFlags.LeaveOpen)) {
                root.CreateStorage("_VBA_PROJECT_CUR"); root.CreateStorage("ObjectPool"); root.Flush(true);
            }
            source = package.ToArray();
        } finally { File.Delete(path); }
        return source;
    }
}
