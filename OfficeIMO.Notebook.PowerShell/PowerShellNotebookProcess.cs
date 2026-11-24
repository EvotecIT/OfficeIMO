using System.ComponentModel;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Processes.Internal;
#if NETFRAMEWORK
using System.Security.AccessControl;
using System.Security.Principal;
#endif

namespace OfficeIMO.Notebook.PowerShell;

/// <summary>Owns the persistent PowerShell process and its private protocol pipe.</summary>
internal sealed class PowerShellNotebookProcess : IDisposable {
    private static readonly Lazy<string> WorkerCommand = new(BuildWorkerCommand);
    private readonly NamedPipeServerStream _pipe;
    private readonly Process _process;
    private readonly StreamReader _reader;
    private int _closed;
    private int _exited;
    private bool _started;
    private OfficeWindowsProcessJob? _windowsJob;
    private bool _assignedToWindowsJob;

    private PowerShellNotebookProcess(string executable, string? workingDirectory) {
        string pipeName = "oin-" + Guid.NewGuid().ToString("N");
        // Unix domain sockets have a short path limit. Use an absolute name when TEMP is deeply nested.
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && Directory.Exists("/tmp")) {
            pipeName = "/tmp/" + pipeName;
        }
        var start = new ProcessStartInfo {
            FileName = executable,
            WorkingDirectory = workingDirectory ?? string.Empty,
            Arguments = WorkerCommand.Value,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        start.EnvironmentVariables["OFFICEIMO_NOTEBOOK_PIPE"] = pipeName;
        _pipe = CreatePipe(pipeName);
        _reader = new StreamReader(_pipe, new UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: false,
            bufferSize: 4096, leaveOpen: true);
        _process = new Process { StartInfo = start, EnableRaisingEvents = true };
        _process.Exited += OnExited;
    }

    internal bool IsOpen => Volatile.Read(ref _closed) == 0 && Volatile.Read(ref _exited) == 0;

    internal string Version { get; private set; } = string.Empty;

    internal static async Task<PowerShellNotebookProcess> StartAsync(string executable, string? workingDirectory,
        TimeSpan startupTimeout, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var worker = new PowerShellNotebookProcess(executable, workingDirectory);
        using var startup = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        startup.CancelAfter(startupTimeout);
        using CancellationTokenRegistration interrupt = worker.InterruptWhenCanceled(startup.Token);
        try {
            try {
                worker._started = worker._process.Start();
            } catch (Win32Exception error) {
                throw new InvalidOperationException("Unable to start PowerShell at '" + executable + "'. Set ExecutablePath to an installed PowerShell executable.", error);
            }
            if (!worker._started) throw new InvalidOperationException("Unable to start PowerShell at '" + executable + "'.");
            worker.AssignWindowsJob();
            worker._process.StandardInput.Close();
            _ = DrainAsync(worker._process.StandardOutput.BaseStream);
            _ = DrainAsync(worker._process.StandardError.BaseStream);
            await worker._pipe.WaitForConnectionAsync(startup.Token).ConfigureAwait(false);
            using JsonDocument handshake = await worker.ReadAsync().ConfigureAwait(false);
            JsonElement response = handshake.RootElement;
            if (GetKind(response) != "ready" || !response.TryGetProperty("text", out JsonElement version)
                || version.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(version.GetString())) {
                throw new WorkerConnectionException("PowerShell did not provide a valid notebook session handshake.");
            }
            startup.CancelAfter(Timeout.Infinite);
            interrupt.Dispose();
            startup.Token.ThrowIfCancellationRequested();
            worker.Version = version.GetString()!;
            return worker;
        } catch (Exception error) {
            worker.Dispose();
            cancellationToken.ThrowIfCancellationRequested();
            if (startup.IsCancellationRequested) {
                throw new TimeoutException("The PowerShell notebook session did not start before its deadline.", error);
            }
            if (error is IOException or ObjectDisposedException) {
                throw new InvalidOperationException("The PowerShell process ended before its notebook session started.", error);
            }
            throw;
        }
    }

    internal CancellationTokenRegistration InterruptWhenCanceled(CancellationToken token) =>
        token.Register(static pipe => ((NamedPipeServerStream)pipe!).Dispose(), _pipe);

    internal async Task SendAsync(string source, int outputCharactersLimit, int outputRecordsLimit, CancellationToken cancellationToken) {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer)) {
            writer.WriteStartObject();
            writer.WriteString("source", source);
            writer.WriteNumber("maxOutputCharacters", outputCharactersLimit);
            writer.WriteNumber("maxOutputRecords", outputRecordsLimit);
            writer.WriteEndObject();
        }
        buffer.WriteByte((byte)'\n');
        try {
            await _pipe.WriteAsync(buffer.GetBuffer(), 0, (int)buffer.Length, cancellationToken).ConfigureAwait(false);
        } catch (Exception error) when (error is IOException or ObjectDisposedException) {
            throw new WorkerConnectionException("The PowerShell worker could not receive the notebook cell.", error);
        }
    }

    internal async Task<JsonDocument> ReadAsync() {
        try {
            string? line = await _reader.ReadLineAsync().ConfigureAwait(false);
            if (line == null) throw new WorkerConnectionException("The PowerShell worker closed its notebook session.");
            return JsonDocument.Parse(line);
        } catch (Exception error) when (error is not WorkerConnectionException && error is IOException or ObjectDisposedException or JsonException or DecoderFallbackException) {
            throw new WorkerConnectionException("The PowerShell worker stopped communicating with the notebook session.", error);
        }
    }

    internal static string GetKind(JsonElement frame) =>
        frame.ValueKind == JsonValueKind.Object && frame.TryGetProperty("kind", out JsonElement kind) &&
        kind.ValueKind == JsonValueKind.String ? kind.GetString()! :
        throw new WorkerConnectionException("The PowerShell worker returned an invalid notebook message.");

    private void AssignWindowsJob() {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        // The bootstrap can connect before this assignment, but no notebook code has been sent yet.
        _windowsJob = OfficeWindowsProcessJob.TryCreateKillOnClose() ??
            throw new InvalidOperationException("Could not create a Windows job for the PowerShell notebook session.");
        _assignedToWindowsJob = _windowsJob.TryAssign(_process.Handle);
        if (!_assignedToWindowsJob) {
            throw new InvalidOperationException("Could not assign the PowerShell notebook worker to its Windows job.");
        }
    }

    public void Dispose() {
        if (Interlocked.Exchange(ref _closed, 1) != 0) return;
        _pipe.Dispose();
        _process.Exited -= OnExited;
        _windowsJob?.Dispose();
        try {
            if (_started && !_process.HasExited && (!_assignedToWindowsJob || !_process.WaitForExit(2_000))) {
#if NET8_0_OR_GREATER
                _process.Kill(entireProcessTree: true);
#else
                _process.Kill();
#endif
                _process.WaitForExit(2_000);
            }
        } catch (Exception error) when (error is Win32Exception or InvalidOperationException or NotSupportedException) {
            // The session is already closed. The process may have exited concurrently.
        } finally {
            _process.Dispose();
        }
    }

    private void OnExited(object? sender, EventArgs args) {
        Volatile.Write(ref _exited, 1);
        _pipe.Dispose();
    }

    private static async Task DrainAsync(Stream stream) {
        try {
            await stream.CopyToAsync(Stream.Null).ConfigureAwait(false);
        } catch (Exception error) when (error is IOException or ObjectDisposedException) {
            // Teardown can close the child descriptors while a read is pending.
        }
    }

    private static string BuildWorkerCommand() {
        using Stream worker = typeof(PowerShellNotebookProcess).Assembly.GetManifestResourceStream("OfficeIMO.Notebook.PowerShell.Worker")
            ?? throw new InvalidOperationException("The PowerShell notebook worker is missing from its assembly.");
        using var reader = new StreamReader(worker, Encoding.UTF8);
        return "-NoLogo -NoProfile -NonInteractive -EncodedCommand " +
            Convert.ToBase64String(Encoding.Unicode.GetBytes(reader.ReadToEnd()));
    }

    private static NamedPipeServerStream CreatePipe(string name) {
#if NET8_0_OR_GREATER
        return new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
#elif NETFRAMEWORK
        using WindowsIdentity identity = WindowsIdentity.GetCurrent();
        SecurityIdentifier user = identity.User ?? throw new InvalidOperationException("Could not identify the notebook session user.");
        var security = new PipeSecurity();
        security.AddAccessRule(new PipeAccessRule(user, PipeAccessRights.FullControl, AccessControlType.Allow));
        return new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous, 0, 0, security);
#else
        // Older reference assemblies do not declare CurrentUserOnly. Require it on the runtime that loads this asset.
        if (!Enum.TryParse("CurrentUserOnly", out PipeOptions currentUserOnly)) {
            throw new PlatformNotSupportedException("Notebook execution requires a runtime that supports current-user named pipes.");
        }
        return new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | currentUserOnly);
#endif
    }

    internal sealed class WorkerConnectionException : IOException {
        internal WorkerConnectionException(string message, Exception? innerException = null) : base(message, innerException) { }
    }
}
