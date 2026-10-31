using System.Diagnostics;
using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private string? _loadedPassword;

    private static WordDocument AssociateEncryptedDocument(WordDocument document, string? path, string password, Stream? source = null, WordLoadOptions? options = null) {
        try {
            if (source != null) OfficeIMO.Core.Internal.OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(source,
                options?.PersistenceMode ?? DocumentPersistenceMode.Explicit, nameof(source));
        } catch { document.Dispose(); throw; }
        document.FilePath = path;
        document._loadedPassword = password;
        document._persistenceMode = options?.PersistenceMode ?? DocumentPersistenceMode.Explicit;
        if (source != null) document.OriginalStream = OfficeIMO.Core.Internal.OfficeDocumentLifecycle.ResolveAssociatedDestination(source, DocumentAccessMode.ReadWrite)!;
        return document;
    }

    /// <summary>Loads a Word file and returns its source-projection findings.</summary>
    public static WordDocument Load(string filePath, WordLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        WordDocument value = Load(filePath, options);
        report = value.CreateLoadReport(filePath, timer.Elapsed);
        return value;
    }

    /// <summary>Loads a Word stream and returns its source-projection findings.</summary>
    public static WordDocument Load(Stream stream, WordLoadOptions? options, out OfficeConversionReport report) {
        var timer = Stopwatch.StartNew();
        WordDocument value = Load(stream, options);
        report = value.CreateLoadReport("Word stream", timer.Elapsed);
        return value;
    }

    internal OfficeConversionReport CreateLoadReport(string source, TimeSpan elapsed) {
        var diagnostics = new List<OfficeDiagnostic>();
        foreach (LegacyDocImportDiagnostic diagnostic in _legacyDocImportDiagnostics) {
            OfficeDiagnosticSeverity severity = diagnostic.Severity switch {
                LegacyDocDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                LegacyDocDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            };
            diagnostics.Add(new OfficeDiagnostic(diagnostic.Code, severity, diagnostic.Message,
                format: "Word", lossKind: LegacyDocLoadResult.ClassifyLoss(diagnostic)));
        }
        foreach (var feature in _legacyDocUnsupportedFeatures)
            diagnostics.Add(new OfficeDiagnostic(feature.Code, OfficeDiagnosticSeverity.Warning, feature.Description,
                feature.EntryPath ?? feature.DetailCode, "Word", OfficeConversionLossKind.Omission));
        foreach (var feature in _legacyDocPreservedFeatures)
            diagnostics.Add(new OfficeDiagnostic(feature.Code, OfficeDiagnosticSeverity.Warning, feature.Description,
                feature.DetailCode, "Word", OfficeConversionLossKind.Omission));
        foreach (var feature in _legacyDocCompoundFeatures)
            diagnostics.Add(new OfficeDiagnostic(feature.Code, OfficeDiagnosticSeverity.Warning, feature.Description,
                feature.EntryPath ?? feature.DetailCode, "Word", OfficeConversionLossKind.Omission));
        return new OfficeConversionReport(source, "Word", itemsWritten: Sections.Count, elapsed: elapsed, diagnostics: diagnostics);
    }

    private string? ResolveSavePassword(WordSaveOptions? options) => options == null ? _loadedPassword : options.ResolvePassword(_loadedPassword);
}
