using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>
/// Keeps one Studio process per user and data root on desktop platforms. A later launch forwards its
/// file arguments to the running instance over a current-user-only local pipe and exits; when the
/// running instance cannot be reached the later launch continues as an independent process.
/// </summary>
internal sealed class StudioSingleInstance : IDisposable {
    internal const int MaximumPaths = 64;
    private const int MaximumMessageBytes = 256 * 1024;
    private const byte Accepted = 1;
    private const byte Rejected = 0;
    private static readonly TimeSpan ReadTimeout = TimeSpan.FromSeconds(5);

    private readonly Mutex? _mutex;
    private readonly CancellationTokenSource _stop = new();
    private Task? _listener;

    private StudioSingleInstance(string name, Mutex? mutex, bool isPrimary) {
        Name = name;
        _mutex = mutex;
        IsPrimary = isPrimary;
    }

    /// <summary>The pipe name shared by every launch with the same user and data root.</summary>
    internal string Name { get; }

    /// <summary>True when this process owns the instance and should receive forwarded files.</summary>
    internal bool IsPrimary { get; }

    internal static string CreateName(string dataRoot) {
        string root = Path.GetFullPath(dataRoot);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()) root = root.ToUpperInvariant();
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(Environment.UserName + "\n" + root));
        // Unix domain socket paths are short; keep the name compact.
        return "OfficeIMO.Studio." + Convert.ToHexString(hash, 0, 8);
    }

    internal static StudioSingleInstance Acquire(string name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        // Unix named mutexes are scoped to the login session unless global; desktop launchers and
        // terminals can run in different sessions, so use the global namespace there. The name
        // already includes the user identity.
        string mutexName = (OperatingSystem.IsWindows() ? string.Empty : @"Global\") + name;
        Mutex? mutex = null;
        try {
            mutex = new Mutex(false, mutexName);
            bool owned;
            try { owned = mutex.WaitOne(0); }
            catch (AbandonedMutexException) { owned = true; }
            if (owned) return new StudioSingleInstance(name, mutex, true);
            mutex.Dispose();
            return new StudioSingleInstance(name, null, false);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException or WaitHandleCannotBeOpenedException or PlatformNotSupportedException) {
            // Without a reliable lock run independently rather than refusing to start.
            mutex?.Dispose();
            return new StudioSingleInstance(name, null, false);
        }
    }

    /// <summary>Sends <paramref name="paths"/> to the running instance. Returns false when it did not accept them.</summary>
    internal bool TryForward(IReadOnlyList<string> paths, TimeSpan timeout) {
        ArgumentNullException.ThrowIfNull(paths);
        if (IsPrimary || paths.Count > MaximumPaths) return false;
        byte[] payload = JsonSerializer.SerializeToUtf8Bytes(new ForwardedLaunch(1, paths.ToArray()));
        if (payload.Length > MaximumMessageBytes) return false;
        // One deadline covers connecting, sending and the reply, so a running instance that stops responding after
        // accepting the connection cannot hold this launch: it falls back to opening its own window.
        using var deadline = new CancellationTokenSource(timeout > TimeSpan.Zero ? timeout : TimeSpan.FromMilliseconds(1));
        try {
            return ForwardAsync(payload, deadline.Token).GetAwaiter().GetResult();
        } catch (Exception error) when (error is IOException or TimeoutException or UnauthorizedAccessException or InvalidOperationException or OperationCanceledException or ArgumentOutOfRangeException) {
            return false;
        }
    }

    private async Task<bool> ForwardAsync(byte[] payload, CancellationToken cancellation) {
        await using var client = new NamedPipeClientStream(".", Name, PipeDirection.InOut, PipeOptions.CurrentUserOnly | PipeOptions.Asynchronous);
        await client.ConnectAsync(cancellation).ConfigureAwait(false);
        AllowForegroundActivation();
        byte[] length = new byte[4];
        System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(length, payload.Length);
        await client.WriteAsync(length, cancellation).ConfigureAwait(false);
        await client.WriteAsync(payload, cancellation).ConfigureAwait(false);
        await client.FlushAsync(cancellation).ConfigureAwait(false);
        byte[] reply = new byte[1];
        return await client.ReadAsync(reply, cancellation).ConfigureAwait(false) == 1 && reply[0] == Accepted;
    }

    /// <summary>
    /// Starts accepting forwarded launches. <paramref name="accept"/> runs on a background thread and must only queue the
    /// paths: it returns true when the running window took them, and the later launch exits only then. Returning false
    /// (or throwing) lets the later launch open its files itself. The listener is free for the next launch as soon as
    /// <paramref name="accept"/> returns, so a slow document open never turns a third launch into a second process.
    /// </summary>
    internal void StartListening(Func<IReadOnlyList<string>, bool> accept) {
        ArgumentNullException.ThrowIfNull(accept);
        if (!IsPrimary || _listener is not null) return;
        _listener = Task.Run(() => ListenAsync(accept, _stop.Token));
    }

    private async Task ListenAsync(Func<IReadOnlyList<string>, bool> accept, CancellationToken cancellation) {
        while (!cancellation.IsCancellationRequested) {
            NamedPipeServerStream server;
            try {
                server = new NamedPipeServerStream(Name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentOutOfRangeException) {
                // An unavailable or overlong Unix pipe path must also permit an independent launch.
                // Another owner holds the name; later launches will fall back to their own process.
                return;
            }
            await using (server) {
                try {
                    await server.WaitForConnectionAsync(cancellation).ConfigureAwait(false);
                    using var read = CancellationTokenSource.CreateLinkedTokenSource(cancellation);
                    read.CancelAfter(ReadTimeout);
                    string[]? paths = await ReadLaunchAsync(server, read.Token).ConfigureAwait(false);
                    if (paths is null) continue;
                    bool accepted;
                    try { accepted = accept(paths); }
                    catch (Exception error) when (error is not OutOfMemoryException) { accepted = false; }
                    server.WriteByte(accepted ? Accepted : Rejected);
                    await server.FlushAsync(cancellation).ConfigureAwait(false);
                } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
                    return;
                } catch (Exception error) when (error is IOException or OperationCanceledException or JsonException or ObjectDisposedException) {
                    // A broken or malformed client must not stop the listener.
                }
            }
        }
    }

    private static async Task<string[]?> ReadLaunchAsync(Stream stream, CancellationToken cancellation) {
        byte[] length = new byte[4];
        await stream.ReadExactlyAsync(length, cancellation).ConfigureAwait(false);
        int size = System.Buffers.Binary.BinaryPrimitives.ReadInt32LittleEndian(length);
        if (size <= 0 || size > MaximumMessageBytes) return null;
        byte[] payload = new byte[size];
        await stream.ReadExactlyAsync(payload, cancellation).ConfigureAwait(false);
        ForwardedLaunch? launch = JsonSerializer.Deserialize<ForwardedLaunch>(payload);
        if (launch is not { Version: 1, Paths: { } paths } || paths.Length > MaximumPaths) return null;
        return paths.Where(static path => !string.IsNullOrWhiteSpace(path) && path.IndexOf('\0') < 0 && Path.IsPathFullyQualified(path)).ToArray();
    }

    private static void AllowForegroundActivation() {
        // The running window may only take focus when the launching process permits it.
        if (OperatingSystem.IsWindows()) _ = AllowSetForegroundWindow(-1);
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AllowSetForegroundWindow(int processId);

    public void Dispose() {
        _stop.Cancel();
        try { _listener?.Wait(TimeSpan.FromSeconds(1)); }
        catch (AggregateException) { }
        if (_mutex is not null) {
            try { _mutex.ReleaseMutex(); }
            catch (ApplicationException) { }
            _mutex.Dispose();
        }
        _stop.Dispose();
    }

    private sealed record ForwardedLaunch(int Version, string[] Paths);
}
