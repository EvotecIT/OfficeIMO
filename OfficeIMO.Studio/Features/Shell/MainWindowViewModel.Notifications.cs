namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private StudioWorkspaceMode _errorWorkspace;
    private StudioDocumentMode _errorDocumentMode;
    private StudioWorkspaceMode _statusWorkspace;
    private StudioDocumentMode _statusDocumentMode;
    private (StudioWorkspaceMode Workspace, StudioDocumentMode Document)? _activeNotificationScope;
    private readonly AsyncLocal<(StudioWorkspaceMode Workspace, StudioDocumentMode Document)?> _notificationContext = new();

    // Async command continuations retain their origin without giving unrelated
    // UI actions ownership of the completed command's notifications.
    private IDisposable BeginNotificationScope() {
        var previous = _notificationContext.Value;
        _notificationContext.Value = previous ?? (WorkspaceMode, DocumentMode);
        return new NotificationScope(() => _notificationContext.Value = previous);
    }

    private sealed class NotificationScope(Action restore) : IDisposable {
        public void Dispose() => restore();
    }

    public bool HasVisibleError => HasError && IsNotificationScope(_errorWorkspace, _errorDocumentMode);

    public bool HasVisibleOperationStatus => HasOperationStatus && !HasVisibleError &&
        IsNotificationScope(_statusWorkspace, _statusDocumentMode);

    private bool IsNotificationScope(StudioWorkspaceMode workspace, StudioDocumentMode documentMode) =>
        WorkspaceMode == workspace && (workspace != StudioWorkspaceMode.PdfWorkspace || DocumentMode == documentMode);

    partial void OnErrorMessageChanged(string? value) {
        (_errorWorkspace, _errorDocumentMode) = _notificationContext.Value ?? _activeNotificationScope ?? (WorkspaceMode, DocumentMode);
        NotifyVisibleNotifications();
    }

    partial void OnOperationStatusChanged(string? value) {
        (_statusWorkspace, _statusDocumentMode) = _notificationContext.Value ?? _activeNotificationScope ?? (WorkspaceMode, DocumentMode);
        NotifyVisibleNotifications();
    }

    partial void OnIsWorkspaceBusyChanged(bool value) {
        if (value) {
            _activeNotificationScope = (WorkspaceMode, DocumentMode);
            OperationStatus = null;
        } else {
            _activeNotificationScope = null;
        }
    }

    private void ResetNotificationScopeForNavigation() {
        if (!IsWorkspaceBusy) _activeNotificationScope = null;
    }

    private void NotifyVisibleNotifications() {
        OnPropertyChanged(nameof(HasVisibleError));
        OnPropertyChanged(nameof(HasVisibleOperationStatus));
    }
}
