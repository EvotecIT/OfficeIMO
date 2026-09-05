using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
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

        Assert.Single(methods, method => method.Name == "ToPdfBytes");
        Assert.DoesNotContain(methods, method => method.Name == "ToPdf");
        Assert.Single(methods, method => method.Name == "ToPdfDocument");
        MethodInfo resultMethod = Assert.Single(methods, method => method.Name == "ToPdfDocumentResult");
        Assert.Equal(typeof(PdfDocumentConversionResult), resultMethod.ReturnType);
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdf" && method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfResult" && method.ReturnType == typeof(PdfSaveResult)));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfAsync" && method.ReturnType == typeof(Task<PdfSaveResult>)));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfResultAsync" && method.ReturnType == typeof(Task<PdfSaveResult>)));
    }

    [Fact]
    public void OdtFacadePreservesProjectionLossAndProducesReadablePdf() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Direct ODT PDF");
        source.AddTrackedParagraphInsertion("Tracked source", "Reviewer");

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        OdfConversionReport projection = Assert.IsType<OdfConversionReport>(
            Assert.Single(result.SourceConversionReports));

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Direct ODT PDF", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(projection.Mappings, mapping =>
            mapping.Feature == "source-tracked-changes" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.True(result.HasLoss);
        OdfConversionLossException loss = Assert.Throws<OdfConversionLossException>(() => result.RequireNoLoss());
        Assert.Same(projection, loss.Report);
        Assert.DoesNotContain(result.Warnings, warning =>
            warning.Code.StartsWith("ODF_", StringComparison.Ordinal));
        Assert.Collection(
            result.ConversionReports,
            report => Assert.Same(projection, report),
            report => Assert.Same(result.Report, report));
    }

    [Fact]
    public void OdsFacadeUsesExcelPdfEngineAndExposesInformationEvidence() {
        OdsDocument source = OdsDocument.Create();
        source.AddSheet("Revenue").Cell(0, 0).SetString("Quarter total");

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        string text = PdfReadDocument.Open(bytes).ExtractText();
        OdfConversionReport projection = Assert.IsType<OdfConversionReport>(
            Assert.Single(result.SourceConversionReports));

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Quarter total", text, StringComparison.Ordinal);
        Assert.Single(PdfReadDocument.Open(bytes).Pages);
        Assert.Contains(projection.Mappings, mapping =>
            mapping.Status == OdfConversionMappingStatus.Converted);
        Assert.DoesNotContain(result.Warnings, warning =>
            warning.Code.StartsWith("ODF_", StringComparison.Ordinal));
    }

    [Fact]
    public void OdpFacadeUsesPowerPointPdfEngineAndKeepsAnimationLoss() {
        OdpPresentation source = OdpPresentation.Create();
        OdpSlide slide = source.AddSlide("Overview");
        slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 8, 2), "Direct ODP PDF");
        OdpRectangle animated = slide.AddRectangle(OdfRect.FromCentimeters(1, 4, 2, 2));
        slide.AddFadeInAnimation(animated, TimeSpan.FromSeconds(1));

        PdfDocumentConversionResult result = source.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();
        string text = PdfReadDocument.Open(bytes).ExtractText();
        OdfConversionReport projection = Assert.IsType<OdfConversionReport>(
            Assert.Single(result.SourceConversionReports));

        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.Contains("Direct ODP PDF", text, StringComparison.Ordinal);
        Assert.Contains(projection.Mappings, mapping =>
            mapping.Feature == "source-presentation-animations" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.True(result.HasLoss);
        OdfConversionLossException loss = Assert.Throws<OdfConversionLossException>(() => result.RequireNoLoss());
        Assert.Same(projection, loss.Report);
    }

    [Fact]
    public async Task DirectFacadeSupportsStreamAndAsyncSaveContracts() {
        OdtDocument source = OdtDocument.Create();
        source.AddParagraph("Stream contract");
        source.AddTrackedParagraphInsertion("Tracked source", "Reviewer");
        using var stream = new MemoryStream();

        PdfSaveResult save = await source.SaveAsPdfAsync(stream);

        Assert.True(save.Succeeded, save.Exception?.Message);
        Assert.True(stream.Length > 0);
        Assert.True(save.HasLoss);
        OdfConversionLossException loss = Assert.Throws<OdfConversionLossException>(() => save.RequireNoLoss());
        Assert.IsType<OdfConversionReport>(loss.Report);
        Assert.Collection(
            save.ConversionReports,
            report => Assert.IsType<OdfConversionReport>(report),
            report => Assert.IsType<PdfConversionReport>(report));
        Assert.All(save.ConversionReports, report => Assert.IsAssignableFrom<IOfficeConversionReport>(report));
    }
}
