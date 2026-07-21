using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.OpenDocument.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentPdfConversionContracts {
    public static IEnumerable<object[]> DirectAdapterTypes() {
        yield return new object[] { typeof(OdtPdfConversionExtensions) };
        yield return new object[] { typeof(OdsPdfConversionExtensions) };
        yield return new object[] { typeof(OdpPdfConversionExtensions) };
    }

    [Theory]
    [MemberData(nameof(DirectAdapterTypes))]
    public void OpenDocumentFacadesExposeTheCanonicalPdfLifecycle(Type adapterType) {
        MethodInfo[] methods = adapterType.GetMethods(BindingFlags.Public | BindingFlags.Static);

        Assert.Single(methods, method => method.Name == "ToPdf");
        Assert.Single(methods, method => method.Name == "ToPdfDocument");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentResult");
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdf" && method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdf" && method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfAsync" && method.ReturnType == typeof(Task<PdfSaveResult>)));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdfAsync" && method.ReturnType == typeof(Task<PdfSaveResult>)));
    }

    [Fact]
    public void OdtFacadePreservesProjectionLossAndProducesReadablePdf() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Direct ODT PDF");
        source.AddTrackedParagraphInsertion("Tracked source", "Reviewer");

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        Assert.Contains("Direct ODT PDF", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "ODF_UNSUPPORTED" &&
            warning.Source.EndsWith(":source-tracked-changes", StringComparison.Ordinal) &&
            warning.Severity == PdfConversionWarningSeverity.Warning &&
            warning.Details["stage"] == "open-document-projection");
    }

    [Fact]
    public void OdsFacadeUsesExcelPdfEngineAndExposesInformationEvidence() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("Revenue").Cell(0, 0).SetString("Quarter total");

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();

        Assert.Contains("Quarter total", text, StringComparison.Ordinal);
        Assert.Single(PdfReadDocument.Open(result.ToBytes()).Pages);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "ODF_CONVERTED" &&
            warning.Severity == PdfConversionWarningSeverity.Information);
    }

    [Fact]
    public void OdpFacadeUsesPowerPointPdfEngineAndKeepsAnimationLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Overview");
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Direct ODP PDF");
        OdpRectangle animated = slide.AddRectangle(OdfRect.FromCentimeters(1, 4, 2, 2));
        slide.AddFadeInAnimation(animated, TimeSpan.FromSeconds(1));

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();

        Assert.Contains("Direct ODP PDF", text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "ODF_UNSUPPORTED" &&
            warning.Source.EndsWith(":source-presentation-animations", StringComparison.Ordinal));
    }

    [Fact]
    public async Task DirectFacadeSupportsStreamAndAsyncSaveContracts() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Stream contract");
        using var stream = new MemoryStream();

        PdfSaveResult save = await source.SaveAsPdfAsync(stream);

        Assert.True(save.Succeeded, save.Exception?.Message);
        Assert.True(stream.Length > 0);
    }
}
