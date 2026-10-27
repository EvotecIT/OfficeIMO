using OfficeIMO.Internal;
using System.Diagnostics;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Drawing;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static bool IsNativePair(OfficeConversionOptions options) {
        OfficeDocumentFormat source = FormatForExtension(Path.GetExtension(options.SourceName));
        return source is OfficeDocumentFormat.Word or OfficeDocumentFormat.Excel or OfficeDocumentFormat.PowerPoint
            && source == FormatForExtension(Path.GetExtension(options.TargetName));
    }

    private static async Task<OfficeConversionReport> ConvertNativeAsync(Stream source, Stream output,
        OfficeConversionOptions options, CancellationToken token, bool analyzeOnly = false) {
        var timer = Stopwatch.StartNew();
        string directory = OfficeIMO.Core.Internal.OfficeTemporaryDirectory.Create("officeimo-native-conversion-");
        try {
            string inputDirectory = Directory.CreateDirectory(Path.Combine(directory, "source")).FullName;
            string outputDirectory = Directory.CreateDirectory(Path.Combine(directory, "target")).FullName;
            string inputPath = Path.Combine(inputDirectory, Path.GetFileName(options.SourceName!));
            string outputPath = Path.Combine(outputDirectory, Path.GetFileName(options.TargetName!));
            using (var captured = new FileStream(inputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                await source.CopyToAsync(captured, 81920, token).ConfigureAwait(false);
            IEnumerable<OfficeConversionDiagnostic> diagnostics;
            try {
                switch (FormatForExtension(Path.GetExtension(inputPath))) {
                    case OfficeDocumentFormat.Word:
                        diagnostics = analyzeOnly
                            ? (await WordDocument.AnalyzeConversionAsync(inputPath, outputPath, WordNativeOptions(options), token).ConfigureAwait(false)).Diagnostics
                            : (await WordDocument.ConvertAsync(inputPath, outputPath, WordNativeOptions(options), token).ConfigureAwait(false)).Report.Diagnostics;
                        break;
                    case OfficeDocumentFormat.Excel:
                        diagnostics = analyzeOnly
                            ? (await ExcelDocument.AnalyzeConversionAsync(inputPath, outputPath, ExcelNativeOptions(options), token).ConfigureAwait(false)).Diagnostics
                            : (await ExcelDocument.ConvertAsync(inputPath, outputPath, ExcelNativeOptions(options), token).ConfigureAwait(false)).Report.Diagnostics;
                        break;
                    case OfficeDocumentFormat.PowerPoint:
                        diagnostics = analyzeOnly
                            ? (await PowerPointPresentation.AnalyzeConversionAsync(inputPath, outputPath, PowerPointNativeOptions(options), token).ConfigureAwait(false)).Diagnostics
                            : (await PowerPointPresentation.ConvertAsync(inputPath, outputPath, PowerPointNativeOptions(options), token).ConfigureAwait(false)).Report.Diagnostics;
                        break;
                    default: throw new NotSupportedException("Choose a supported native Office format pair.");
                }
            } catch (WordDocumentConversionException error) {
                throw new OfficeConversionException(error.Message, NativeReport(options, error.Result.Report.Diagnostics, timer.Elapsed), error);
            } catch (ExcelDocumentConversionException error) {
                throw new OfficeConversionException(error.Message, NativeReport(options, error.Result.Report.Diagnostics, timer.Elapsed), error);
            } catch (PowerPointPresentationConversionException error) {
                throw new OfficeConversionException(error.Message, NativeReport(options, error.Result.Report.Diagnostics, timer.Elapsed), error);
            }
            var report = NativeReport(options, diagnostics, timer.Elapsed, itemsWritten: analyzeOnly ? 0 : 1);
            if (options.FailOnLoss) report.ThrowIfLoss();
            if (analyzeOnly) return report;
            using var artifact = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (artifact.Length > options.MaximumOutputBytes) throw new InvalidDataException("The artifact exceeds the conversion output limit.");
            await artifact.CopyToAsync(output, 81920, token).ConfigureAwait(false);
            return report;
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static OfficeConversionReport NativeReport(OfficeConversionOptions options,
        IEnumerable<OfficeConversionDiagnostic> diagnostics, TimeSpan elapsed, long itemsWritten = 0) =>
        new(options.SourceName!, options.TargetName!, itemsWritten: itemsWritten, elapsed: elapsed, diagnostics: diagnostics.Select(d => new OfficeDiagnostic(
            d.Code, d.Severity switch {
                OfficeConversionDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                OfficeConversionDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            }, d.Message, d.SourceLocation, FormatForExtension(Path.GetExtension(options.SourceName)).ToString(), d.LossKind)));

    private static WordDocumentConversionOptions WordNativeOptions(OfficeConversionOptions settings) => new() {
        LoadOptions = WordInputOptions(settings), SaveOptions = settings.WordSave,
        CompatibilityMode = settings.CompatibilityMode, EmbedSourceWhenLossy = settings.EmbedSourceWhenLossy,
        LossPolicy = settings.FailOnLoss ? OfficeConversionLossPolicy.Block : OfficeConversionLossPolicy.Allow,
        VisualFallbackRenderer = document => document.ExportImages(OfficeImageExportFormat.Png)
    };

    private static ExcelDocumentConversionOptions ExcelNativeOptions(OfficeConversionOptions settings) => new() {
        LoadOptions = ExcelInputOptions(settings), SaveOptions = settings.ExcelSave,
        CompatibilityMode = settings.CompatibilityMode, EmbedSourceWhenLossy = settings.EmbedSourceWhenLossy,
        LossPolicy = settings.FailOnLoss ? OfficeConversionLossPolicy.Block : OfficeConversionLossPolicy.Allow,
        VisualFallbackMaxColumns = settings.VisualFallbackMaxColumns, VisualFallbackMaxRows = settings.VisualFallbackMaxRows
    };

    private static PowerPointPresentationConversionOptions PowerPointNativeOptions(OfficeConversionOptions settings) => new() {
        LoadOptions = PowerPointInputOptions(settings), SaveOptions = settings.PowerPointSave,
        CompatibilityMode = settings.CompatibilityMode, EmbedSourceWhenLossy = settings.EmbedSourceWhenLossy,
        LossPolicy = settings.FailOnLoss ? OfficeConversionLossPolicy.Block : OfficeConversionLossPolicy.Allow
    };

    private static WordLoadOptions WordInputOptions(OfficeConversionOptions settings) {
        WordLoadOptions options = settings.WordLoad?.Snapshot() ?? new();
        options.PersistenceMode = DocumentPersistenceMode.Explicit;
        options.Password ??= settings.SourcePassword;
        options.MaxInputBytes = Math.Min(options.MaxInputBytes, settings.MaximumInputBytes);
        options.PackageSecurity = LimitPackage(options.PackageSecurity, settings);
        options.OpenSettings = LimitXml(options.OpenSettings, settings);
        options.ReportUnsupportedContent = true;
        return options;
    }

    private static ExcelLoadOptions ExcelInputOptions(OfficeConversionOptions settings) {
        ExcelLoadOptions options = settings.ExcelLoad?.Snapshot() ?? new();
        options.PersistenceMode = DocumentPersistenceMode.Explicit;
        options.Password ??= settings.SourcePassword;
        options.MaxInputBytes = Math.Min(options.MaxInputBytes ?? long.MaxValue, settings.MaximumInputBytes);
        options.PackageSecurity = LimitPackage(options.PackageSecurity, settings);
        options.OpenSettings = LimitXml(options.OpenSettings, settings);
        options.ReportUnsupportedContent = true;
        return options;
    }

    private static PowerPointLoadOptions PowerPointInputOptions(OfficeConversionOptions settings) {
        PowerPointLoadOptions options = settings.PowerPointLoad?.Snapshot() ?? new();
        options.PersistenceMode = DocumentPersistenceMode.Explicit;
        options.Password ??= settings.SourcePassword;
        options.MaxInputBytes = Math.Min(options.MaxInputBytes ?? long.MaxValue, settings.MaximumInputBytes);
        options.PackageSecurity = LimitPackage(options.PackageSecurity, settings);
        options.OpenSettings = LimitXml(options.OpenSettings, settings);
        options.ReportUnsupportedContent = true;
        return options;
    }

    private static OfficePackageSecurityOptions LimitPackage(OfficePackageSecurityOptions? policy, OfficeConversionOptions settings) {
        var copy = policy?.Snapshot() ?? OfficePackageSecurityOptions.SecureDefaults;
        copy.MaxPackageBytes = Math.Min(copy.MaxPackageBytes, settings.MaximumInputBytes);
        copy.MaxXmlCharactersInPart = Math.Min(copy.MaxXmlCharactersInPart, settings.MaximumXmlCharactersInPart);
        return copy;
    }

    private static OfficeOpenXmlLoadSettings LimitXml(OfficeOpenXmlLoadSettings? options, OfficeConversionOptions settings) {
        var copy = options?.Snapshot() ?? new();
        copy.MaxCharactersInPart = Math.Min(copy.MaxCharactersInPart <= 0 ? long.MaxValue : copy.MaxCharactersInPart, settings.MaximumXmlCharactersInPart);
        return copy;
    }
}
