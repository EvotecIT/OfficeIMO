using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;

internal static class ExampleExecution {
    public static Task<IntelligenceXOfficeAiExecutor> ConnectAsync(ExampleOptions options, bool images, CancellationToken cancellationToken) =>
        IntelligenceXOfficeAiExecutor.ConnectAsync(new OfficeAiExecutionProfile {
            Id = options.Local ? "local-example" : "hosted-example", Provider = options.Copilot ? "GitHubCopilot" : options.Endpoint is null ? "ChatGPT" : "CompatibleHttp",
            Model = options.Model, MaxRequestCharacters = options.RequestCharacters, SupportsImages = images && !options.Copilot, EnforcesJsonSchema = !options.PromptedJson && !options.Copilot, IsLocal = options.Local
        }, new OfficeAiIntelligenceXOptions {
            Transport = options.Copilot ? OfficeAiIntelligenceXTransport.CopilotCli : options.Endpoint is null ? OfficeAiIntelligenceXTransport.ChatGpt : OfficeAiIntelligenceXTransport.CompatibleHttp,
            Endpoint = options.Endpoint, ApiKey = options.Endpoint is null && !options.Copilot ? null : Environment.GetEnvironmentVariable("OFFICEIMO_AI_API_KEY"),
            PreferCurrentCodexSession = options.CodexSession
        }, cancellationToken);
}
