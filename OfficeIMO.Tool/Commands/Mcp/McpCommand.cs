using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Mcp;
using OfficeIMO.Ocr;
using OfficeIMO.Tool.Commands.Pdf;

namespace OfficeIMO.Tool.Commands.Mcp;

internal static class McpCommand {
    internal const string Usage = """
OfficeIMO.Tool - Model Context Protocol

Usage:
  officeimo mcp serve --stdio [--ocr-provider-assembly <trusted-provider.dll>]
             [--ocr-option <key=value>]

Optional OCR provider assemblies and executable/model paths are trusted server-startup configuration.
PDF tools cannot load assemblies or choose executable paths. No OCR provider is installed automatically.
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
        if (args.Length < 2 ||
            !args[0].Equals("serve", StringComparison.OrdinalIgnoreCase) ||
            !args[1].Equals("--stdio", StringComparison.OrdinalIgnoreCase)) {
            await standardError.WriteLineAsync("MCP requires 'serve --stdio'.").ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }

        try {
            var ocrCatalog = new OcrEngineCatalog();
            var assemblies = new List<string>();
            var providerOptions = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int index = 2; index < args.Length; index++) {
                switch (args[index]) {
                    case "--ocr-provider-assembly": assemblies.Add(PdfWorkflowArguments.Value(args, ref index, args[index])); break;
                    case "--ocr-option": PdfWorkflowArguments.AddProviderOption(providerOptions, PdfWorkflowArguments.Value(args, ref index, args[index])); break;
                    default: throw new AgentUsageException("Unknown MCP startup option.");
                }
            }
            PdfOcrProviderLoader.LoadExplicitAssemblies(ocrCatalog, assemblies);
            HostApplicationBuilder builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings {
                Args = Array.Empty<string>()
            });
            builder.Logging.ClearProviders();
            builder.Services.AddSingleton(new OfficeImoAgentService(
                AgentPathPolicy.FromMcpEnvironment(), pdfOcrCatalog: ocrCatalog, pdfOcrProviderOptions: providerOptions));
            builder.Services.AddSingleton(serviceProvider => new OfficeImoMcpTools(
                serviceProvider.GetRequiredService<OfficeImoAgentService>()));
            var serializerOptions = AgentJson.CreateSerializerOptions();
            builder.Services
                .AddMcpServer(options => {
                    options.ServerInfo = new Implementation {
                        Name = "officeimo",
                        Title = "OfficeIMO local documents and mailboxes",
                        Version = typeof(McpCommand).Assembly.GetName().Version?.ToString(3) ?? "unknown",
                        Description = "Bounded local inspection, search, fetch, conversion, and explicit PDF output workflows."
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
        } catch (AgentUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (Exception exception) {
            await standardError.WriteLineAsync(
                "MCP server failed: " + exception.GetType().Name)
                .ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }
}
