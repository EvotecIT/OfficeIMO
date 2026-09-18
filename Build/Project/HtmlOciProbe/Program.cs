using System.Diagnostics;
using System.Text.Json.Nodes;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;

if (args.Length > 0 && new[] { "info", "create", "inspect", "start", "rm", "ps" }.Contains(args[0], StringComparer.Ordinal)) {
    await ProxyPodmanAsync();
    return;
}

if (args.Length != 2) throw new ArgumentException("Usage: OfficeIMO.Html.OciProbe <full-sha256-image-id> <deployed-worker-dll>");
string imageId = args[0];
string workerPath = Path.GetFullPath(args[1]);
var provider = HtmlRuntimeWorkerManifest.Load(workerPath);

await RunCaseAsync("protocol", "<p id='result'>Isolated worker</p>", false);
await RunCaseAsync("cancellation", "<p id='result'>Started</p>", true);
await RunRejectedInspectionAsync(false);
await RunRejectedInspectionAsync(true);
await RunRemovalRetryAsync();
Console.WriteLine("OCI worker protocol, cancellation, post-create cleanup and removal retry passed.");

async Task RunCaseAsync(string name, string html, bool cancel) {
    using HtmlOciWorkerLease lease = await HtmlOciWorkerLease.StartAsync("podman", imageId, CancellationToken.None);
    string containerName = lease.ContainerName;
    var request = new HtmlScriptRequest { Html = html, Timeout = TimeSpan.FromSeconds(15), SessionTimeout = TimeSpan.FromSeconds(30) }.Snapshot();
    await using (var session = new HtmlProcessRuntimeSession(lease, AngleSharpDomServices.Instance, request,
        Guid.NewGuid().ToString("N"), Guid.NewGuid().ToString("N"), provider, new HtmlRuntimeTraceOptions(), null)) {
        await session.OpenAsync(CancellationToken.None);
        if (cancel) {
            using var stop = new CancellationTokenSource(TimeSpan.FromMilliseconds(200));
            try {
                await session.ExecuteAsync("while (true) {}", stop.Token);
                throw new InvalidOperationException("The runaway script was not cancelled.");
            } catch (OperationCanceledException) when (stop.IsCancellationRequested) { }
        } else {
            HtmlScriptCapture capture = await session.CaptureAsync();
            if (capture.Document.QuerySelector("#result")?.TextContent != "Isolated worker")
                throw new InvalidOperationException("The isolated worker returned the wrong document.");
        }
    }
    using var exists = new Process { StartInfo = new ProcessStartInfo("podman") {
        UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
    } };
    exists.StartInfo.ArgumentList.Add("container");
    exists.StartInfo.ArgumentList.Add("exists");
    exists.StartInfo.ArgumentList.Add(containerName);
    if (!exists.Start()) throw new InvalidOperationException("Podman inspection did not start.");
    await exists.WaitForExitAsync();
    if (exists.ExitCode != 1) throw new InvalidOperationException($"The {name} container was not removed: {containerName}.");
    Console.WriteLine(name + ": container removed");
}

async Task RunRejectedInspectionAsync(bool failFirstRemoval) {
    string[] before = await ManagedContainersAsync();
    string proxyExecutable = Path.Combine(AppContext.BaseDirectory, "OfficeIMO.Html.OciProbe");
    if (!File.Exists(proxyExecutable)) throw new InvalidOperationException("The Linux probe apphost is required for the inspection proxy.");
    string marker = Path.Combine(Path.GetTempPath(), "officeimo-oci-rejected-" + Guid.NewGuid().ToString("N"));
    Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MODE", failFirstRemoval ? "reject-inspection-fail-first-removal" : "reject-inspection");
    if (failFirstRemoval) Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MARKER", marker);
    try {
        try {
            await using HtmlOciWorkerLease unexpected = await HtmlOciWorkerLease.StartAsync(proxyExecutable, imageId, CancellationToken.None);
            throw new InvalidOperationException("The rejected inspection was admitted.");
        } catch (HtmlOciWorkerStartException error) when (
            error.InnerException is HtmlScriptRuntimeException { Message: "The container engine did not retain the required isolation policy." }
            && error.ContainerRemoved && error.CleanupError == null) { }
        if (failFirstRemoval && !File.Exists(marker)) throw new InvalidOperationException("The first startup removal did not fail as expected.");
    } finally {
        Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MODE", null);
        Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MARKER", null);
        if (File.Exists(marker)) File.Delete(marker);
    }
    string[] after = await ManagedContainersAsync();
    if (!before.SequenceEqual(after, StringComparer.Ordinal))
        throw new InvalidOperationException("A rejected inspection left an OfficeIMO container behind.");
    Console.WriteLine(failFirstRemoval ? "rejected inspection: first removal failed, retry removed container"
        : "rejected inspection: created container removed");
}

async Task RunRemovalRetryAsync() {
    string[] before = await ManagedContainersAsync();
    string proxyExecutable = Path.Combine(AppContext.BaseDirectory, "OfficeIMO.Html.OciProbe");
    string marker = Path.Combine(Path.GetTempPath(), "officeimo-oci-removal-" + Guid.NewGuid().ToString("N"));
    Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MODE", "fail-first-removal");
    Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MARKER", marker);
    try {
        await using (HtmlOciWorkerLease lease = await HtmlOciWorkerLease.StartAsync(proxyExecutable, imageId, CancellationToken.None)) {
            await WaitForRunningAsync(lease.ContainerName);
            lease.Stop();
            await lease.WaitForRemovalAsync();
        }
        if (!File.Exists(marker)) throw new InvalidOperationException("The first removal did not fail as expected.");
    } finally {
        Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MODE", null);
        Environment.SetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MARKER", null);
        if (File.Exists(marker)) File.Delete(marker);
    }
    string[] after = await ManagedContainersAsync();
    if (!before.SequenceEqual(after, StringComparer.Ordinal))
        throw new InvalidOperationException("A failed first removal left an OfficeIMO container behind.");
    Console.WriteLine("removal retry: exact container removed");
}

async Task WaitForRunningAsync(string name) {
    for (int attempt = 0; attempt < 20; attempt++) {
        using var inspect = new Process { StartInfo = new ProcessStartInfo("podman") {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
        } };
        foreach (string arg in new[] { "inspect", "--format", "{{.State.Running}}", name }) inspect.StartInfo.ArgumentList.Add(arg);
        if (!inspect.Start()) throw new InvalidOperationException("Podman state inspection did not start.");
        string state = await inspect.StandardOutput.ReadToEndAsync();
        await inspect.WaitForExitAsync();
        if (inspect.ExitCode == 0 && state.Trim() == "true") return;
        await Task.Delay(250);
    }
    throw new InvalidOperationException("The worker did not reach a running container state.");
}

async Task<string[]> ManagedContainersAsync() {
    using var list = new Process { StartInfo = new ProcessStartInfo("podman") {
        UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true
    } };
    foreach (string arg in new[] { "ps", "--all", "--filter", "name=officeimo-html-", "--format", "{{.Names}}" })
        list.StartInfo.ArgumentList.Add(arg);
    if (!list.Start()) throw new InvalidOperationException("Podman listing did not start.");
    string output = await list.StandardOutput.ReadToEndAsync();
    await list.WaitForExitAsync();
    if (list.ExitCode != 0) throw new InvalidOperationException("Podman listing failed.");
    return output.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Order(StringComparer.Ordinal).ToArray();
}

async Task ProxyPodmanAsync() {
    string? mode = Environment.GetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MODE");
    if (mode is "fail-first-removal" or "reject-inspection-fail-first-removal" && args[0] == "rm") {
        string marker = Environment.GetEnvironmentVariable("OFFICEIMO_OCI_PROBE_MARKER")
            ?? throw new InvalidOperationException("The removal marker is missing.");
        try {
            using (new FileStream(marker, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { }
            Console.Error.Write("Injected first removal failure.");
            Environment.ExitCode = 125;
            return;
        } catch (IOException) when (File.Exists(marker)) { }
    }
    using var command = new Process { StartInfo = new ProcessStartInfo("podman") {
        UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = args[0] == "start",
        RedirectStandardOutput = true, RedirectStandardError = true
    } };
    foreach (string arg in args) command.StartInfo.ArgumentList.Add(arg);
    if (!command.Start()) throw new InvalidOperationException("The Podman proxy could not start.");
    Task<string> output = command.StandardOutput.ReadToEndAsync();
    Task<string> error = command.StandardError.ReadToEndAsync();
    await command.WaitForExitAsync();
    string body = await output;
    if (command.ExitCode == 0 && args[0] == "inspect" && mode is "reject-inspection" or "reject-inspection-fail-first-removal") {
        JsonNode inspection = JsonNode.Parse(body) ?? throw new InvalidOperationException("Podman inspection was empty.");
        inspection[0]!["HostConfig"]!["NetworkMode"] = "bridge";
        body = inspection.ToJsonString();
    }
    Console.Out.Write(body);
    Console.Error.Write(await error);
    Environment.ExitCode = command.ExitCode;
}
