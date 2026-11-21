using System.Diagnostics;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace OfficeIMO.Tool.Tests;

/// <summary>
/// Runs the stdio MCP server as a child process the test owns and connects a client to it.
/// </summary>
/// <remarks>
/// Disposing a <see cref="StdioClientTransport"/> client does not close the server's standard input, so the server keeps
/// running until the transport's shutdown timeout expires and the process is killed. The kill is not awaited, so the
/// server can still be alive, and still hold its working directory, when the test deletes that directory. This host
/// ends the session the way the stdio transport expects a client to: it closes standard input, waits for the server to
/// exit, and only kills it if it does not exit in time. Disposal returns after the process has exited.
/// </remarks>
internal sealed class McpTestServer : IAsyncDisposable {
    private static readonly TimeSpan GracefulExitTimeout = TimeSpan.FromSeconds(30);
    private readonly Process _process;

    private McpTestServer(Process process) => _process = process;

    internal McpClient Client { get; private set; } = null!;

    /// <summary>Starts the server described by <paramref name="options"/> and completes the MCP handshake.</summary>
    internal static async Task<McpTestServer> StartAsync(StdioClientTransportOptions options, CancellationToken cancellationToken) {
        var startInfo = new ProcessStartInfo(options.Command) {
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            WorkingDirectory = options.WorkingDirectory ?? string.Empty
        };
        foreach (string argument in options.Arguments ?? Array.Empty<string>()) startInfo.ArgumentList.Add(argument);
        if (options.EnvironmentVariables is not null) {
            foreach (KeyValuePair<string, string?> variable in options.EnvironmentVariables) {
                if (variable.Value is null) startInfo.Environment.Remove(variable.Key);
                else startInfo.Environment[variable.Key] = variable.Value;
            }
        }

        var process = new Process { StartInfo = startInfo };
        var server = new McpTestServer(process);
        // Standard error is drained so a chatty server can never block on a full pipe.
        process.ErrorDataReceived += (_, _) => { };
        process.Start();
        process.BeginErrorReadLine();
        try {
            var transport = new StreamClientTransport(process.StandardInput.BaseStream, process.StandardOutput.BaseStream);
            server.Client = await McpClient.CreateAsync(transport, cancellationToken: cancellationToken).ConfigureAwait(false);
            return server;
        } catch {
            await server.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    public async ValueTask DisposeAsync() {
        try {
            if (Client is not null) await Client.DisposeAsync().ConfigureAwait(false);
        } finally {
            try {
                _process.StandardInput.Close();
            } catch (Exception exception) when (exception is IOException or ObjectDisposedException) {
                // The transport already closed the stream; the server sees end of input either way.
            }

            using (var graceful = new CancellationTokenSource(GracefulExitTimeout)) {
                try {
                    await _process.WaitForExitAsync(graceful.Token).ConfigureAwait(false);
                } catch (OperationCanceledException) {
                    _process.Kill(entireProcessTree: true);
                    await _process.WaitForExitAsync().ConfigureAwait(false);
                }
            }
            _process.Dispose();
        }
    }
}
