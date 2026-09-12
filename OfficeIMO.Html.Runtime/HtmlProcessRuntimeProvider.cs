using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html.Runtime;

/// <summary>Runs one trusted scripted document in a disposable worker process with a deadline and bounded response.</summary>
/// <remarks>Process separation permits termination of runaway scripts. It is not an OS sandbox for hostile code.</remarks>
public sealed class HtmlProcessRuntimeProvider : IHtmlScriptRuntimeProvider {
    private readonly string _workerPath;
    private readonly string _dotnetExecutable;
    private readonly IHtmlDomServices _services;

    /// <summary>Uses a deployed OfficeIMO runtime worker DLL and the chosen inert DOM query/serialization services.</summary>
    public HtmlProcessRuntimeProvider(string workerAssemblyPath, IHtmlDomServices domServices, string dotnetExecutable = "dotnet") {
        _workerPath = Path.GetFullPath(workerAssemblyPath ?? throw new ArgumentNullException(nameof(workerAssemblyPath)));
        if (!File.Exists(_workerPath)) throw new FileNotFoundException("The runtime worker is not deployed.", _workerPath);
        _services = domServices ?? throw new ArgumentNullException(nameof(domServices));
        _dotnetExecutable = dotnetExecutable ?? throw new ArgumentNullException(nameof(dotnetExecutable));
    }

    /// <inheritdoc />
    public async Task<HtmlScriptCapture> CaptureTrustedAsync(HtmlScriptRequest request, CancellationToken cancellationToken = default) {
        HtmlScriptRequest input = (request ?? throw new ArgumentNullException(nameof(request))).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        using var deadline = new CancellationTokenSource(input.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        var start = new ProcessStartInfo(_dotnetExecutable) { UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
        start.ArgumentList.Add(_workerPath);
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new HtmlScriptRuntimeException("The runtime worker could not start.");
        using var stop = operation.Token.Register(() => Kill(process));
        Task<string> output = ReadOrStopAsync(process, process.StandardOutput, input.MaxOutputCharacters, operation.Token);
        Task<string> errors = ReadOrStopAsync(process, process.StandardError, 64 * 1024, operation.Token);
        try {
            string json = JsonSerializer.Serialize(input);
            await process.StandardInput.WriteAsync(json.AsMemory(), operation.Token).ConfigureAwait(false);
            process.StandardInput.Close();
            await process.WaitForExitAsync(operation.Token).ConfigureAwait(false);
            string response = await output.ConfigureAwait(false);
            string stderr = await errors.ConfigureAwait(false);
            if (process.ExitCode != 0) throw new HtmlScriptRuntimeException("The runtime worker failed: " + stderr);
            HtmlRuntimeWireDocument wire = JsonSerializer.Deserialize<HtmlRuntimeWireDocument>(response)
                ?? throw new HtmlScriptRuntimeException("The runtime worker returned no capture.");
            if (wire.Error != null) throw new HtmlScriptRuntimeException(wire.Error);
            return new HtmlScriptCapture(wire.Materialize(_services, input, operation.Token), wire.ProviderId);
        } catch (Exception) when (operation.IsCancellationRequested) {
            cancellationToken.ThrowIfCancellationRequested();
            throw new TimeoutException("The scripted document did not complete within its execution deadline.");
        } finally {
            Kill(process);
            await process.WaitForExitAsync(CancellationToken.None).ConfigureAwait(false);
            operation.Cancel();
            try { await Task.WhenAll(output, errors).ConfigureAwait(false); } catch { /* Observe stream completion after termination. */ }
        }
    }

    private static void Kill(Process process) {
        try { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        catch (InvalidOperationException) { /* Already exited. */ }
    }

    private static async Task<string> ReadOrStopAsync(Process process, TextReader reader, int maximum, CancellationToken token) {
        try { return await ReadBoundedAsync(reader, maximum, token).ConfigureAwait(false); }
        catch { Kill(process); throw; }
    }

    internal static async Task<string> ReadBoundedAsync(TextReader reader, int maximum, CancellationToken cancellationToken) {
        var result = new StringBuilder();
        var buffer = new char[4096];
        int read;
        while ((read = await reader.ReadAsync(buffer.AsMemory(), cancellationToken).ConfigureAwait(false)) != 0) {
            if ((long)result.Length + read > maximum) throw new HtmlScriptRuntimeException("The worker response exceeded its character budget.");
            result.Append(buffer, 0, read);
        }
        return result.ToString();
    }
}
