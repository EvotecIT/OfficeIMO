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
            PdfSaveResult result = path
                ? useAsync ? await saves.PathAsync(outputPath) : saves.Path(outputPath)
                : useAsync ? await saves.StreamAsync(output) : saves.Stream(output);

            Assert.False(result.Succeeded);
            OdfConversionLossException loss = Assert.IsType<OdfConversionLossException>(result.Exception);
            Assert.True(result.HasLoss);
            Assert.Collection(result.ConversionReports,
                report => Assert.Equal(loss.Report.FidelityDiagnostics.Select(d => (d.Code, d.LossKind, d.Source, d.Location)),
                    Assert.IsType<OdfConversionReport>(report).FidelityDiagnostics.Select(d => (d.Code, d.LossKind, d.Source, d.Location))),
                report => Assert.Same(result.Report, Assert.IsType<PdfConversionReport>(report)));
            Assert.Contains(loss.Report.Mappings, mapping => mapping.Feature == expectedFeature &&
                mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Contains(result.FidelityDiagnostics, diagnostic => diagnostic.Location == expectedFeature &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission);
            Assert.Equal(0L, result.BytesWritten);
            Assert.Equal(original, path ? File.ReadAllBytes(outputPath) : output.ToArray());
            Assert.True(output.CanWrite);
        } finally {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    private static (
        Func<string, PdfSaveResult> Path,
        Func<Stream, PdfSaveResult> Stream,
        Func<string, Task<PdfSaveResult>> PathAsync,
        Func<Stream, Task<PdfSaveResult>> StreamAsync) CreateSavers(string format) {
        switch (format) {
            case "odt": {
                // Existing fixture: UnsupportedOdtWritingModesAreReportedAndEnforced.
                OdtDocument source = OdtDocument.Create();
                source.AddParagraph("Vertical").WritingMode = "tb-rl";
                var options = new WordOpenDocumentConversionOptions {
                    LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
                };
                return (
                    path => source.SaveAsPdfResult(path, options),
                    stream => source.SaveAsPdfResult(stream, options),
                    path => source.SaveAsPdfResultAsync(path, options),
                    stream => source.SaveAsPdfResultAsync(stream, options));
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
                    path => source.SaveAsPdfResult(path, options),
                    stream => source.SaveAsPdfResult(stream, options),
                    path => source.SaveAsPdfResultAsync(path, options),
                    stream => source.SaveAsPdfResultAsync(stream, options));
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
                    path => persisted.SaveAsPdfResult(path, options),
                    stream => persisted.SaveAsPdfResult(stream, options),
                    path => persisted.SaveAsPdfResultAsync(path, options),
                    stream => persisted.SaveAsPdfResultAsync(stream, options));
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }
}
