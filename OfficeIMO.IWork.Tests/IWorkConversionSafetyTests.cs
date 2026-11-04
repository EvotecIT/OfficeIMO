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
            using var result = source.ToWordResult(preview); report = result.Report;
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelResult(preview); report = result.Report;
        } else {
            using var result = source.ToPowerPointResult(preview); report = result.Report;
        }
        Assert.Equal(IWorkProjectionKind.VisualFallback, report.ProjectionKind);
        Assert.False(report.HasCompleteVisualCoverage);
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_VISUAL_FALLBACK"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
    }

    private static void ConvertValue(IWorkSourceDocument source, IWorkConversionOptions options) {
        using IDisposable value = source.Kind switch {
            IWorkDocumentKind.Pages => source.ToWordResult(options).RequireCompleteEditableReconstruction(),
            IWorkDocumentKind.Numbers => source.ToExcelResult(options).RequireCompleteEditableReconstruction(),
            _ => source.ToPowerPointResult(options).RequireCompleteEditableReconstruction()
        };
    }

    private static void ConvertValue(Stream stream, IWorkDocumentKind kind, IWorkConversionOptions options) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages).ToWordResult(options).RequireCompleteEditableReconstruction(),
            IWorkDocumentKind.Numbers => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers).ToExcelResult(options).RequireCompleteEditableReconstruction(),
            _ => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote).ToPowerPointResult(options).RequireCompleteEditableReconstruction()
        };
    }

    private static void ConvertValue(string path, IWorkDocumentKind kind, IWorkConversionOptions options) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages).ToWordResult(options).RequireCompleteEditableReconstruction(),
            IWorkDocumentKind.Numbers => IWorkSourceDocument.Open(path, IWorkDocumentKind.Numbers).ToExcelResult(options).RequireCompleteEditableReconstruction(),
            _ => IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote).ToPowerPointResult(options).RequireCompleteEditableReconstruction()
        };
    }

    private static void ConvertValue(Stream stream, IWorkDocumentKind kind, IWorkConversionOptions options, CancellationToken token) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Pages, null, token).ToWordResult(options, cancellationToken: token).RequireCompleteEditableReconstruction(),
            IWorkDocumentKind.Numbers => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers, null, token).ToExcelResult(options, cancellationToken: token).RequireCompleteEditableReconstruction(),
            _ => IWorkSourceDocument.Open(stream, IWorkDocumentKind.Keynote, null, token).ToPowerPointResult(options, cancellationToken: token).RequireCompleteEditableReconstruction()
        };
    }

    private static void ConvertValue(string path, IWorkDocumentKind kind, IWorkConversionOptions options, CancellationToken token) {
        using IDisposable value = kind switch {
            IWorkDocumentKind.Pages => IWorkSourceDocument.Open(path, IWorkDocumentKind.Pages, null, token).ToWordResult(options, cancellationToken: token).RequireCompleteEditableReconstruction(),
            IWorkDocumentKind.Numbers => IWorkSourceDocument.Open(path, IWorkDocumentKind.Numbers, null, token).ToExcelResult(options, cancellationToken: token).RequireCompleteEditableReconstruction(),
            _ => IWorkSourceDocument.Open(path, IWorkDocumentKind.Keynote, null, token).ToPowerPointResult(options, cancellationToken: token).RequireCompleteEditableReconstruction()
        };
    }
}
