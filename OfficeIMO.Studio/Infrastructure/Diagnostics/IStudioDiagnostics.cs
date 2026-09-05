namespace OfficeIMO.Studio.Infrastructure.Diagnostics;

internal enum StudioDiagnosticLevel {
    Information,
    Warning,
    Error,
    Critical
}

/// <summary>Records privacy-bounded operational evidence without document content.</summary>
internal interface IStudioDiagnostics {
    string DirectoryPath { get; }

    void Write(StudioDiagnosticLevel level, string area, string code, Exception? exception = null);

    StudioSupportSnapshot CreateSupportSnapshot();
}

internal sealed record StudioSupportSnapshot(
    string Product,
    string Version,
    string OperatingSystem,
    string Runtime,
    string Architecture,
    string UiCulture,
    string DiagnosticsDirectory,
    string PrivacyNotice);
