namespace OfficeIMO.Bibliography.Tests;

internal static class BibliographyCancellationTest {
    internal static void AssertObserved(Action<CancellationToken> operation, int delayMilliseconds = 1) {
        using var cancellation = new CancellationTokenSource();
        using var started = new ManualResetEventSlim();
        var cancellationThread = new Thread(() => {
            started.Wait();
            Thread.Sleep(delayMilliseconds);
            cancellation.Cancel();
        }) { IsBackground = true };
        cancellationThread.Start();

        try {
            started.Set();
            Assert.Throws<OperationCanceledException>(() => operation(cancellation.Token));
        } finally {
            cancellationThread.Join();
        }
    }
}
