namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    /// <summary>The document the window shows; a document can open in another tab, so a start tool continues there.</summary>
    internal Func<MainWindowViewModel>? PresentedDocument { get; set; }

    /// <summary>
    /// Runs a start-screen tool. A tool that works on a document, chosen with none open, asks for a PDF first and then
    /// runs on the document that opens; cancelling the file choice leaves everything as it was.
    /// </summary>
    internal async Task StartToolAsync(string id) {
        StudioCommandItem command = Commands[id];
        if (command.IsAvailable) {
            await command.ExecuteAsync().ConfigureAwait(true);
            return;
        }
        if (!command.RequiresDocument || HasDocument || !CanStartDocumentTransition) return;
        await OpenCommand.ExecuteAsync(null).ConfigureAwait(true);
        MainWindowViewModel opened = PresentedDocument?.Invoke() ?? this;
        if (opened.HasDocument && opened.Commands[id] is { IsAvailable: true } next) await next.ExecuteAsync().ConfigureAwait(true);
    }
}
