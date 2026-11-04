using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentPdfStrictFailureSiblingTests {
    [Theory]
    [InlineData("odt", false, false)]
    [InlineData("odt", false, true)]
    [InlineData("odt", true, false)]
    [InlineData("odt", true, true)]
    [InlineData("ods", false, false)]
    [InlineData("ods", false, true)]
    [InlineData("ods", true, false)]
    [InlineData("ods", true, true)]
    [InlineData("odp", false, false)]
    [InlineData("odp", false, true)]
    [InlineData("odp", true, false)]
    [InlineData("odp", true, true)]
    public async Task StrictSaveFailuresRetainSourceLossEvidence(string format, bool useAsync, bool path) {
        var saves = CreateSavers(format);
        string expectedFeature = format switch {
            "odt" => "writing-mode",
            "ods" => "validations",
            "odp" => "hyperlink-target-behavior",
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
        byte[] original = { 1, 2, 3 };
        string outputPath = Path.Combine(Path.GetTempPath(),
            "officeimo-" + format + "-strict-" + Guid.NewGuid().ToString("N") + ".pdf");
        using var output = new MemoryStream(original.ToArray(), writable: true);
        try {
            if (path) File.WriteAllBytes(outputPath, original);
            OfficeConversionException exception = path
                ? useAsync ? await Assert.ThrowsAnyAsync<OfficeConversionException>(() => saves.PathAsync(outputPath)) : Assert.ThrowsAny<OfficeConversionException>(() => saves.Path(outputPath))
                : useAsync ? await Assert.ThrowsAnyAsync<OfficeConversionException>(() => saves.StreamAsync(output)) : Assert.ThrowsAny<OfficeConversionException>(() => saves.Stream(output));

            OdfConversionLossException loss = Assert.IsType<OdfConversionLossException>(exception);
            Assert.True(loss.Report.HasLoss);
            Assert.Contains(loss.Report.Mappings, mapping => mapping.Feature == expectedFeature &&
                mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Contains(loss.Report.Diagnostics, diagnostic => diagnostic.Location == expectedFeature &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Equal(original, path ? File.ReadAllBytes(outputPath) : output.ToArray());
            Assert.True(output.CanWrite);
        } finally {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    private static (
        Func<string, OfficeConversionReport> Path,
        Func<Stream, OfficeConversionReport> Stream,
        Func<string, Task<OfficeConversionReport>> PathAsync,
        Func<Stream, Task<OfficeConversionReport>> StreamAsync) CreateSavers(string format) {
        switch (format) {
            case "odt": {
                // Existing fixture: UnsupportedOdtWritingModesAreReportedAndEnforced.
                OdtDocument source = OdtDocument.Create();
                source.AddParagraph("Vertical").WritingMode = "tb-rl";
                var options = new WordOpenDocumentConversionOptions {
                    LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
                };
                return (
                    path => source.SaveAsPdf(path, Export(options)),
                    stream => source.SaveAsPdf(stream, Export(options)),
                    path => source.SaveAsPdfAsync(path, Export(options)),
                    stream => source.SaveAsPdfAsync(stream, Export(options)));
            }
            case "ods": {
                // Existing fixture: OversizedInlineOdsValidationIsReportedInsteadOfWritingAnInvalidExcelFormula.
                OdsDocument source = OdsDocument.Create();
                OdsValidation validation = source.AddValidation("Oversized",
                    OdsValidationConditionSyntax.CreateList(new[] { new string('x', 254) }));
                source.AddSheet("Data").Cell(0, 0).ValidationName = validation.Name;
                var options = new ExcelOpenDocumentConversionOptions {
                    LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
                };
                return (
                    path => source.SaveAsPdf(path, Export(options)),
                    stream => source.SaveAsPdf(stream, Export(options)),
                    path => source.SaveAsPdfAsync(path, Export(options)),
                    stream => source.SaveAsPdfAsync(stream, Export(options)));
            }
            case "odp": {
                // Existing fixture: OdpHyperlinkTargetBehaviorIsReportedWhileTheTargetIsPreserved.
                OdpPresentation source = OdpPresentation.Create();
                OdpHyperlink hyperlink = source.AddSlide("Source")
                    .AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2))
                    .AddParagraph()
                    .AddHyperlink("Docs", "https://example.test/docs");
                hyperlink.TargetFrameName = "_blank";
                hyperlink.ShowBehavior = "new";
                using var input = new MemoryStream(source.ToBytes());
                OdpPresentation persisted = OdpPresentation.Load(input);
                var options = new PowerPointOpenDocumentConversionOptions {
                    LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
                };
                return (
                    path => persisted.SaveAsPdf(path, Export(options)),
                    stream => persisted.SaveAsPdf(stream, Export(options)),
                    path => persisted.SaveAsPdfAsync(path, Export(options)),
                    stream => persisted.SaveAsPdfAsync(stream, Export(options)));
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }

    private static OdtPdfExportOptions Export(WordOpenDocumentConversionOptions options) => new() { ConversionOptions = options };
    private static OdsPdfExportOptions Export(ExcelOpenDocumentConversionOptions options) => new() { ConversionOptions = options };
    private static OdpPdfExportOptions Export(PowerPointOpenDocumentConversionOptions options) => new() { ConversionOptions = options };
}
