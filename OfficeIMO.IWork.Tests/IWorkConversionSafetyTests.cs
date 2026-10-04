using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "pages")]
    [InlineData(IWorkDocumentKind.Numbers, "numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "key")]
    public void Value_only_entrypoints_reject_partial_editable_output_even_when_explicitly_enabled(
        IWorkDocumentKind kind, string extension) {
        using var package = ParagraphLayoutPackage(kind, BytesField(13, FloatField(2, float.NaN)));
        var options = new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        Assert.Throws<InvalidOperationException>(() => ConvertValue(source, options));
        package.Position = 0;
        Assert.Throws<InvalidOperationException>(() => ConvertValue(package, kind, options));
        package.Position = 0;
        Assert.Throws<InvalidOperationException>(() => ConvertValue(package, kind, options, CancellationToken.None));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
        try {
            File.WriteAllBytes(path, package.ToArray());
            Assert.Throws<InvalidOperationException>(() => ConvertValue(path, kind, options));
            Assert.Throws<InvalidOperationException>(() => ConvertValue(path, kind, options, CancellationToken.None));
        } finally { File.Delete(path); }
        package.Position = 0;
        Assert.True(ConvertUnitReport(package, kind, visual: false).IsPartialEditableReconstruction);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages", IWorkDocumentKind.Pages)]
    [InlineData("nim-iwork/simple.numbers", IWorkDocumentKind.Numbers)]
    [InlineData("nim-iwork/simple.key", IWorkDocumentKind.Keynote)]
    public void Visual_previews_require_explicit_opt_in_and_value_only_entrypoints_still_reject_them(
        string fixture, IWorkDocumentKind kind) {
        IWorkSourceDocument source = IWorkSourceDocument.Open(CorpusFixture(fixture));
        var strict = new IWorkConversionOptions { Mode = IWorkConversionMode.VisualOnly };
        Assert.Throws<InvalidDataException>(() => ConvertValue(source, strict));
        var preview = new IWorkConversionOptions {
            Mode = IWorkConversionMode.VisualOnly, RequireCompleteVisualCoverage = false
        };
        Assert.Throws<InvalidOperationException>(() => ConvertValue(source, preview));
        IWorkConversionReport report;
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(preview); report = result.Report;
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(preview); report = result.Report;
        } else {
            using var result = source.ToPowerPointPresentationResult(preview); report = result.Report;
        }
        Assert.Equal(IWorkProjectionKind.VisualFallback, report.ProjectionKind);
        Assert.False(report.HasCompleteVisualCoverage);
        Assert.Contains(report.FidelityDiagnostics, d => d.Code == "IWORK_VISUAL_FALLBACK"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
    }

    private static void ConvertValue(IWorkSourceDocument source, IWorkConversionOptions options) {
        using IDisposable value = source.Kind switch {
            IWorkDocumentKind.Pages => source.ToWordDocument(options),
            IWorkDocumentKind.Numbers => source.ToExcelDocument(options),
            _ => source.ToPowerPointPresentation(options)
        };
    }

    private static void ConvertValue(Stream stream, IWorkDocumentKind kind, IWorkConversionOptions options) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWord(stream, conversionOptions: options),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcel(stream, conversionOptions: options),
            _ => PowerPointIWorkConverter.ConvertKeynoteToPowerPoint(stream, conversionOptions: options)
        };
    }

    private static void ConvertValue(string path, IWorkDocumentKind kind, IWorkConversionOptions options) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWord(path, conversionOptions: options),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcel(path, conversionOptions: options),
            _ => PowerPointIWorkConverter.ConvertKeynoteToPowerPoint(path, conversionOptions: options)
        };
    }

    private static void ConvertValue(Stream stream, IWorkDocumentKind kind, IWorkConversionOptions options, CancellationToken token) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWord(stream, null, options, token),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcel(stream, null, options, token),
            _ => PowerPointIWorkConverter.ConvertKeynoteToPowerPoint(stream, null, options, token)
        };
    }

    private static void ConvertValue(string path, IWorkDocumentKind kind, IWorkConversionOptions options, CancellationToken token) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => WordIWorkConverter.ConvertPagesToWord(path, null, options, token),
            IWorkDocumentKind.Numbers => ExcelIWorkConverter.ConvertNumbersToExcel(path, null, options, token),
            _ => PowerPointIWorkConverter.ConvertKeynoteToPowerPoint(path, null, options, token)
        };
    }
}
