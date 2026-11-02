namespace OfficeIMO.Studio.Tests;

internal sealed class PausedPdfPublicationGuard : IOfficePublicationGuard {
    internal TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
    internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

    public async Task<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
        if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase)) return true;
        Started.TrySetResult();
        await Release.Task.WaitAsync(cancellationToken);
        return true;
    }
}
