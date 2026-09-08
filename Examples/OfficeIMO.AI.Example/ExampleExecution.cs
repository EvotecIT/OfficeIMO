using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;

internal static class ExampleExecution {
    internal static OfficeAiExecutionProfile BuildProfile(ExampleOptions options, bool images) {
        if (images && !options.ModelSupportsImages)
            throw new ArgumentException("Image input requires --model-supports-images for a model known to accept images.");
        return new OfficeAiExecutionProfile {
            Id = options.Local ? "local-example" : "hosted-example", Provider = options.Copilot ? "GitHubCopilot" : options.Endpoint is null ? "ChatGPT" : "CompatibleHttp",
            Model = options.Model, MaxRequestCharacters = options.RequestCharacters, SupportsImages = options.ModelSupportsImages, EnforcesJsonSchema = !options.PromptedJson, IsLocal = options.Local
        };
    }

    public static Task<IntelligenceXOfficeAiExecutor> ConnectAsync(ExampleOptions options, bool images, CancellationToken cancellationToken) =>
        IntelligenceXOfficeAiExecutor.ConnectAsync(BuildProfile(options, images), new OfficeAiIntelligenceXOptions {
            Transport = options.Copilot ? OfficeAiIntelligenceXTransport.CopilotNative : options.Endpoint is null ? OfficeAiIntelligenceXTransport.ChatGpt : OfficeAiIntelligenceXTransport.CompatibleHttp,
            Endpoint = options.Endpoint, ApiKey = options.Endpoint is null && !options.Copilot ? null : Environment.GetEnvironmentVariable("OFFICEIMO_AI_API_KEY"),
            PreferCurrentCodexSession = options.CodexSession
        }, cancellationToken);
}
