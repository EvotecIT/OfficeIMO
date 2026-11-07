using OfficeIMO.Drawing;
using OfficeIMO.Internal;
using OfficeIMO.Excel.Legacy;
using OfficeIMO.Excel.LegacyXls;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    internal OfficeLegacyImportReport? EarlyBiffImportReport { get; private set; }
    private WeakReference<Stream>? _earlyBiffSourceStream;

    private void RememberEarlyBiffSource(Stream? source) {
        if (EarlyBiffImportReport == null || source == null) return;
        _earlyBiffSourceStream = new WeakReference<Stream>(source);
        if (source is FileStream file && Path.IsPathRooted(file.Name)) {
            _legacyXlsSourcePath = Path.GetFullPath(file.Name);
        }
    }

    private void EnsureEarlyBiffOutputPath(string path) {
        if (EarlyBiffImportReport == null) return;
        if (IsEarlyBiffSourcePath(path)) throw new NotSupportedException("The original early BIFF file cannot be replaced.");
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException("An early BIFF import can only be saved to a separate XLSX destination.");
        }
    }

    private bool IsEarlyBiffSourcePath(string path) =>
        _legacyXlsSourcePath is string source && OfficePathIdentity.AreEquivalent(source, path);

    private void EnsureEarlyBiffOutputStream(Stream destination, ExcelFileFormat format) {
        if (EarlyBiffImportReport == null) return;
        if (format != ExcelFileFormat.Xlsx) {
            throw new NotSupportedException("An early BIFF import can only be saved as XLSX.");
        }
        if ((_earlyBiffSourceStream?.TryGetTarget(out Stream? source) == true && ReferenceEquals(source, destination))
            || (destination is FileStream file && Path.IsPathRooted(file.Name) && IsEarlyBiffSourcePath(file.Name))) {
            throw new NotSupportedException("The original early BIFF stream cannot be replaced.");
        }
    }

    private static void ValidateEarlyBiffSecurity(OfficeLegacyImportReport? report, OfficePackageSecurityOptions? policy) {
        if (report == null || policy == null) return;
        OfficeLegacyInertContentKind content = report.InertContent;
        if ((content & OfficeLegacyInertContentKind.Macros) != 0 && policy.Macros == OfficePackageContentPolicy.Reject) {
            throw new OfficePackageSecurityException(OfficePackageSecurityRule.Macros, "The early BIFF source contains an Excel macro sheet.");
        }
        if ((content & OfficeLegacyInertContentKind.EmbeddedObjects) != 0 && policy.EmbeddedPayloads == OfficePackageContentPolicy.Reject) {
            throw new OfficePackageSecurityException(OfficePackageSecurityRule.EmbeddedPayloads, "The early BIFF source contains an embedded object record.");
        }
        if ((content & OfficeLegacyInertContentKind.ExternalLinks) != 0 && policy.ExternalRelationships == OfficePackageContentPolicy.Reject) {
            throw new OfficePackageSecurityException(OfficePackageSecurityRule.ExternalRelationships, "The early BIFF source contains an external reference record.");
        }
    }

    private static ExcelDocument LoadEarlyBiff(byte[] bytes, bool readOnly, string? path,
        LegacyXlsImportOptions options, OfficePackageSecurityOptions? securityPolicy) {
        var limits = new OfficeLegacyImportLimits();
        limits.MaxInputBytes = Math.Min(limits.MaxInputBytes, options.MaxInputBytes);
        (LegacySpreadsheetModel model, string profile) = EarlyBiffWorkbook.Parse(bytes, limits, options.CancellationToken);
        var report = new OfficeLegacyImportReport(profile, model.Quality, model.Findings, model.InertContent, model.RecoveredCellCount);
        ValidateEarlyBiffSecurity(report, securityPolicy);
        ExcelDocument document = LegacySpreadsheetImporter.Project(model, options.CancellationToken);
        try {
            if (readOnly) document = ReopenProjectedWorkbookReadOnly(document);
            options.CancellationToken.ThrowIfCancellationRequested();
            document.SourceFormat = ExcelFileFormat.Xls;
            document._legacyXlsSourcePath = path == null ? null : Path.GetFullPath(path);
            document.EarlyBiffImportReport = report;
            document._sourceLoadDiagnostics = report.Diagnostics.ToArray();
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }
}
