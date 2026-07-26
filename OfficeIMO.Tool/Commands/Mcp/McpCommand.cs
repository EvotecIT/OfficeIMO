using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Mcp;

namespace OfficeIMO.Tool.Commands.Mcp;

internal static class McpCommand {
    internal const string Usage = """
OfficeIMO.Tool - Model Context Protocol

Usage:
  officeimo mcp serve --stdio
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(args);
        if (args.Length == 1 && args[0] is "help" or "--help" or "-h") {
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (args.Length != 2 ||
            !args[0].Equals("serve", StringComparison.OrdinalIgnoreCase) ||
            !args[1].Equals("--stdio", StringComparison.OrdinalIgnoreCase)) {
            await standardError.WriteLineAsync("MCP requires 'serve --stdio'.").ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }

        try {
            HostApplicationBuilder builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings {
                Args = Array.Empty<string>()
            });
            builder.Logging.ClearProviders();
            builder.Services.AddSingleton(new OfficeImoAgentService());
            builder.Services.AddSingleton(serviceProvider => new OfficeImoMcpTools(
                serviceProvider.GetRequiredService<OfficeImoAgentService>()));
            var serializerOptions = AgentJson.CreateSerializerOptions();
            builder.Services
                .AddMcpServer(options => {
                    options.ServerInfo = new Implementation {
                        Name = "officeimo",
                        Title = "OfficeIMO local documents and mailboxes",
                        Version = typeof(McpCommand).Assembly.GetName().Version?.ToString(3) ?? "unknown",
                        Description = "Bounded local inspection, search, fetch, and conversion."
                    };
                    options.ServerInstructions = OfficeImoMcpTools.ServerInstructions;
                })
                .WithStdioServerTransport()
                .WithTools<OfficeImoMcpTools>(serializerOptions);
            using IHost host = builder.Build();
            await host.RunAsync(cancellationToken).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (Exception exception) {
            await standardError.WriteLineAsync(
                "MCP server failed: " + exception.GetType().Name + ": " + exception.Message)
                .ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }
}