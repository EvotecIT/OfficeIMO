using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

// An internal transport building block. Admission of public content requires a
// separate provider, verified isolation report and resource acquisition broker.
internal sealed class HtmlOciWorkerLease : IDisposable, IAsyncDisposable {
    private readonly string _podmanExecutable;
    private readonly string _containerName;
    private readonly object _sync = new();
    private Task? _removal;
    private int _disposed;

    private HtmlOciWorkerLease(string podmanExecutable, string containerName, Process attach) {
        _podmanExecutable = podmanExecutable;
        _containerName = containerName;
        Process = attach;
    }

    internal Process Process { get; }
    internal string ContainerName => _containerName;

    internal static async Task<HtmlOciWorkerLease> StartAsync(string podmanExecutable, string imageId, CancellationToken token) {
        ArgumentException.ThrowIfNullOrWhiteSpace(podmanExecutable);
        ValidateImageId(imageId);
        string name = "officeimo-html-" + Guid.NewGuid().ToString("N");
        bool createAttempted = false;
        // The name is chosen before create, so a cancelled or failed create can
        // still be cleaned up without parsing its possibly incomplete output.
        try {
            string engineInfo = await RunAsync(podmanExecutable,
                new[] { "info", "--format", "json" }, TimeSpan.FromSeconds(15), token).ConfigureAwait(false);
            string[] baselineCapabilities = VerifyEngine(engineInfo);
            createAttempted = true;
            await RunAsync(podmanExecutable, new[] { "create", "--interactive", "--name", name,
                "--network", "none", "--read-only", "--pids-limit", "32", "--memory", "512m",
                "--cpus", "1", "--cap-drop", "all", "--security-opt", "no-new-privileges",
                "--user", "65532:65532", imageId }, TimeSpan.FromSeconds(30), token).ConfigureAwait(false);
            string inspection = await RunAsync(podmanExecutable, new[] { "inspect", name }, TimeSpan.FromSeconds(15), token).ConfigureAwait(false);
            VerifyInspection(inspection, name, imageId, baselineCapabilities);
            token.ThrowIfCancellationRequested();
            var start = new ProcessStartInfo(podmanExecutable) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true, StandardErrorEncoding = Encoding.UTF8
            };
            foreach (string arg in new[] { "start", "--attach", "--interactive", name }) start.ArgumentList.Add(arg);
            var attach = new Process { StartInfo = start };
            bool started = false;
            try {
                if (!attach.Start()) throw new HtmlScriptRuntimeException("The isolated worker could not attach.");
                started = true;
                token.ThrowIfCancellationRequested();
                return new HtmlOciWorkerLease(podmanExecutable, name, attach);
            } catch {
                if (started && !attach.HasExited) attach.Kill(entireProcessTree: true);
                attach.Dispose();
                throw;
            }
        } catch {
            if (createAttempted) await RemoveWithRetryAsync(podmanExecutable, name).ConfigureAwait(false);
            throw;
        }
    }

    internal void Stop() {
        try { if (!Process.HasExited) Process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { }
        catch (Win32Exception) { /* Removal below is the authoritative container stop. */ }
        lock (_sync) _removal ??= RemoveAsync(_podmanExecutable, _containerName);
    }

    internal async Task WaitForRemovalAsync() {
        for (int attempt = 0; attempt < 2; attempt++) {
            Task removal;
            lock (_sync) removal = _removal ??= RemoveAsync(_podmanExecutable, _containerName);
            try { await removal.ConfigureAwait(false); return; }
            catch when (attempt == 0) {
                lock (_sync) if (ReferenceEquals(_removal, removal)) _removal = null;
            }
        }
    }

    public void Dispose() {
        if (Volatile.Read(ref _disposed) != 0) return;
        Stop();
        WaitForRemovalAsync().GetAwaiter().GetResult();
        if (Interlocked.Exchange(ref _disposed, 1) == 0) Process.Dispose();
    }

    public async ValueTask DisposeAsync() {
        if (Volatile.Read(ref _disposed) != 0) return;
        Stop();
        await WaitForRemovalAsync().ConfigureAwait(false);
        if (Interlocked.Exchange(ref _disposed, 1) == 0) Process.Dispose();
    }

    private static void ValidateImageId(string imageId) {
        ArgumentException.ThrowIfNullOrWhiteSpace(imageId);
        if (!imageId.StartsWith("sha256:", StringComparison.Ordinal) || imageId.Length != 71 ||
            !imageId.AsSpan(7).ToString().All(Uri.IsHexDigit))
            throw new ArgumentException("The isolated worker image must use a full immutable sha256 image ID.", nameof(imageId));
    }

    private static async Task RemoveAsync(string executable, string name) {
        await RunAsync(executable, new[] { "rm", "--force", "--ignore", name }, TimeSpan.FromSeconds(15), CancellationToken.None).ConfigureAwait(false);
        string remaining = await RunAsync(executable, new[] { "ps", "--all", "--filter", "name=" + name,
            "--format", "{{.Names}}" }, TimeSpan.FromSeconds(15), CancellationToken.None).ConfigureAwait(false);
        if (remaining.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Contains(name, StringComparer.Ordinal))
            throw new HtmlScriptRuntimeException("The isolated worker container was not removed: " + name);
    }

    private static async Task RemoveWithRetryAsync(string executable, string name) {
        try { await RemoveAsync(executable, name).ConfigureAwait(false); }
        catch {
            try { await RemoveAsync(executable, name).ConfigureAwait(false); }
            catch (Exception error) {
                throw new HtmlScriptRuntimeException("The isolated worker container could not be removed after startup failure: "
                    + name + ". " + error.Message);
            }
        }
    }

    private static string[] VerifyEngine(string json) {
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement host = document.RootElement.GetProperty("host");
        JsonElement security = host.GetProperty("security");
        if (!security.GetProperty("rootless").GetBoolean() || !security.GetProperty("seccompEnabled").GetBoolean())
            throw new NotSupportedException("The isolated worker requires rootless Podman with seccomp enabled.");
        var controllers = host.GetProperty("cgroupControllers").EnumerateArray()
            .Select(value => value.GetString()).ToHashSet(StringComparer.Ordinal);
        if (!controllers.IsSupersetOf(new[] { "cpu", "memory", "pids" }))
            throw new NotSupportedException("The isolated worker requires CPU, memory and PID cgroup controllers.");
        string capabilities = security.GetProperty("capabilities").GetString() ?? string.Empty;
        string[] result = capabilities.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (result.Length == 0 || result.Any(capability => !capability.StartsWith("CAP_", StringComparison.Ordinal)))
            throw new NotSupportedException("The container engine did not report its baseline capabilities.");
        return result;
    }

    private static void VerifyInspection(string json, string name, string imageId, IReadOnlyList<string> baselineCapabilities) {
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Array || root.GetArrayLength() != 1)
            throw new HtmlScriptRuntimeException("The container engine returned an invalid isolation inspection.");
        JsonElement item = root[0];
        JsonElement host = item.GetProperty("HostConfig");
        JsonElement config = item.GetProperty("Config");
        bool admitted = item.GetProperty("Name").GetString() == name
            && item.GetProperty("Image").GetString() == imageId[7..]
            && host.GetProperty("Memory").GetInt64() == 512L * 1024 * 1024
            && host.GetProperty("PidsLimit").GetInt32() == 32
            && host.GetProperty("NanoCpus").GetInt64() == 1_000_000_000L
            && host.GetProperty("ReadonlyRootfs").GetBoolean()
            && host.GetProperty("NetworkMode").GetString() == "none"
            && !host.GetProperty("Privileged").GetBoolean()
            && host.GetProperty("PidMode").GetString() == "private"
            && host.GetProperty("Binds").GetArrayLength() == 0
            && item.GetProperty("Mounts").GetArrayLength() == 0
            && config.GetProperty("User").GetString() == "65532:65532";
        JsonElement security = host.GetProperty("SecurityOpt");
        admitted &= security.ValueKind == JsonValueKind.Array && security.GetArrayLength() == 1
            && security[0].GetString() == "no-new-privileges";
        JsonElement addedCapabilities = host.GetProperty("CapAdd");
        JsonElement droppedCapabilities = host.GetProperty("CapDrop");
        admitted &= addedCapabilities.ValueKind == JsonValueKind.Array && addedCapabilities.GetArrayLength() == 0
            && droppedCapabilities.ValueKind == JsonValueKind.Array
            && droppedCapabilities.EnumerateArray().Select(value => value.GetString())
                .ToHashSet(StringComparer.Ordinal).IsSupersetOf(baselineCapabilities);
        if (!admitted) throw new HtmlScriptRuntimeException("The container engine did not retain the required isolation policy.");
    }

    private static async Task<string> RunAsync(string executable, IReadOnlyList<string> args, TimeSpan timeout, CancellationToken token) {
        var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string arg in args) start.ArgumentList.Add(arg);
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new HtmlScriptRuntimeException("The container engine did not start.");
        using var deadline = new CancellationTokenSource(timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token, deadline.Token);
        Task<string> stdout = process.StandardOutput.ReadToEndAsync(operation.Token);
        Task<string> stderr = process.StandardError.ReadToEndAsync(operation.Token);
        try {
            await process.WaitForExitAsync(operation.Token).ConfigureAwait(false);
            string error = await stderr.ConfigureAwait(false);
            string output = await stdout.ConfigureAwait(false);
            if (process.ExitCode != 0) throw new HtmlScriptRuntimeException("The container engine command failed: " + TrimError(error));
            if (output.Length > 64 * 1024) throw new HtmlScriptRuntimeException("The container engine command exceeded its output budget.");
            return output;
        } catch (OperationCanceledException) {
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
            catch (InvalidOperationException) { }
            await process.WaitForExitAsync(CancellationToken.None).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            throw new TimeoutException("The container engine command exceeded its deadline.");
        }
    }

    private static string TrimError(string error) => error.Length > 1024 ? error[..1024] : error.Trim();
}
