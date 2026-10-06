using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Host preferences, messages and persistence services consumed by restart-session operations.</summary>
internal interface IStudioSessionEnvironment : IDisposable {
    StudioSessionStore SessionStore { get; }
    PdfWorkspaceRecoveryStore Recovery { get; }
    StudioDocumentStorage Storage { get; }
    bool RememberSession { get; }
    event EventHandler? PreferencesChanged;
    event EventHandler? SessionCleared;
    string Text(string key);
    string Format(string key, int count);
    void ReportFailure(string code, Exception error);
}
