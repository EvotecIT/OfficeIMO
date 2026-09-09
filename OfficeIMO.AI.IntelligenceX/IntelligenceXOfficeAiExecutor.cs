using IntelligenceX.Json;
using IntelligenceX.OpenAI;
using IntelligenceX.OpenAI.AppServer.Models;
using IntelligenceX.OpenAI.CompatibleHttp;
using IntelligenceX.Treatment;

namespace OfficeIMO.AI.IntelligenceX;

/// <summary>Thin adapter to IX Treatment. Each request starts fresh and supplies only inline evidence.</summary>
public sealed class IntelligenceXOfficeAiExecutor : IOfficeAiExecutor, IDisposable {
    private readonly IntelligenceXClient? _client;
    private readonly ITreatmentProvider _provider;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _disposed;

    private IntelligenceXOfficeAiExecutor(IntelligenceXClient client, OfficeAiExecutionProfile profile) {
        _client = client; _provider = new OpenAIChatTreatmentProvider(client); Profile = profile;
    }

    /// <inheritdoc />
    public OfficeAiExecutionProfile Profile { get; }

    /// <summary>Connects a restricted SDK client. Local profiles require an explicit loopback-compatible endpoint.</summary>
    public static async Task<IntelligenceXOfficeAiExecutor> ConnectAsync(OfficeAiExecutionProfile profile,
        OfficeAiIntelligenceXOptions? connection = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(profile);
        connection ??= new();
        profile.Validate();
        var options = new IntelligenceXClientOptions { DefaultModel = profile.Model, EnableUsageTelemetry = false };
        switch (connection.Transport) {
            case OfficeAiIntelligenceXTransport.CopilotNative:
                if (profile.IsLocal || connection.Endpoint is not null || connection.PreferCurrentCodexSession)
                    throw new ArgumentException("Copilot uses its hosted inference service and GitHub credentials.", nameof(connection));
                if (connection.CopilotOptions is not null && connection.ApiKey is not null)
                    throw new ArgumentException("Configure the GitHub credential in CopilotOptions when those options are supplied.", nameof(connection));
                options.TransportKind = OpenAITransportKind.CopilotNative;
                options.CopilotOptions = connection.CopilotOptions ?? new() { GitHubToken = connection.ApiKey, Streaming = connection.Streaming };
                break;
            case OfficeAiIntelligenceXTransport.ChatGpt:
                if (profile.IsLocal || connection.Endpoint is not null || connection.ApiKey is not null)
                    throw new ArgumentException("ChatGPT uses the IX auth store and cannot be labelled local or redirected.", nameof(connection));
                options.TransportKind = OpenAITransportKind.Native;
                options.NativeOptions.PreferCurrentCodexSession = connection.PreferCurrentCodexSession;
                options.NativeOptions.LoadCodexAuthJson = connection.LoadCodexAuthJson || connection.PreferCurrentCodexSession;
                if (connection.AuthStore is not null) options.NativeOptions.AuthStore = connection.AuthStore;
                if (connection.AuthStore is not null || connection.AccountId is not null) options.NativeOptions.AuthAccountId = connection.AccountId;
                options.NativeOptions.PersistCodexAuthJson = false;
                options.NativeOptions.EnableModelFallback = false;
                options.NativeOptions.EnableToolSchemaFallback = false;
                options.NativeOptions.AllowSensitiveDiagnostics = false;
                options.NativeOptions.ImageGeneration = new() { Enabled = false };
                break;
            case OfficeAiIntelligenceXTransport.CompatibleHttp:
                Uri endpoint = connection.Endpoint ?? throw new ArgumentException("CompatibleHttp requires an explicit endpoint.", nameof(connection));
                if (!endpoint.IsAbsoluteUri || endpoint.Scheme is not ("https" or "http") || endpoint.UserInfo.Length > 0
                    || endpoint.Query.Length > 0 || endpoint.Fragment.Length > 0)
                    throw new ArgumentException("The endpoint must be an absolute HTTP(S) base URL without credentials, query or fragment.", nameof(connection));
                if (profile.IsLocal && !endpoint.IsLoopback) throw new ArgumentException("Local profiles require a loopback endpoint.", nameof(profile));
                if (endpoint.Scheme == "http" && !endpoint.IsLoopback) throw new ArgumentException("Non-loopback endpoints require HTTPS.", nameof(connection));
                options.TransportKind = OpenAITransportKind.CompatibleHttp;
                options.CompatibleHttpOptions.BaseUrl = endpoint.AbsoluteUri;
                options.CompatibleHttpOptions.ApiKey = connection.ApiKey;
                options.CompatibleHttpOptions.AuthMode = string.IsNullOrEmpty(connection.ApiKey) ? OpenAICompatibleHttpAuthMode.None : OpenAICompatibleHttpAuthMode.Bearer;
                options.CompatibleHttpOptions.AllowInsecureHttp = endpoint.IsLoopback;
                options.CompatibleHttpOptions.Streaming = connection.Streaming;
                options.CompatibleHttpOptions.AllowAutoRedirect = false;
                options.CompatibleHttpOptions.UseProxy = !profile.IsLocal;
                break;
            default: throw new ArgumentOutOfRangeException(nameof(connection));
        }
        IntelligenceXClient client = await IntelligenceXClient.ConnectAsync(options, cancellationToken).ConfigureAwait(false);
        return new IntelligenceXOfficeAiExecutor(client, profile with { });
    }

    /// <summary>Lists models available through this connection without sending document evidence.</summary>
    public Task<ModelListResult> ListModelsAsync(CancellationToken cancellationToken = default) =>
        WithClientAsync(client => client.ListModelsAsync(cancellationToken), cancellationToken);

    /// <summary>Reads the selected account without starting an inference request.</summary>
    public Task<AccountInfo> GetAccountAsync(CancellationToken cancellationToken = default) =>
        WithClientAsync(client => client.GetAccountAsync(cancellationToken), cancellationToken);

    /// <summary>Signs in through the native ChatGPT browser flow; no agent CLI is launched.</summary>
    public Task LoginChatGptAsync(Action<string> onUrl, CancellationToken cancellationToken = default) =>
        WithClientAsync(async client => {
            await client.LoginChatGptAndWaitAsync(onUrl: onUrl, useLocalListener: true,
                timeout: TimeSpan.FromMinutes(5), cancellationToken: cancellationToken).ConfigureAwait(false);
            return true;
        }, cancellationToken);

    /// <summary>Signs in through the native GitHub device flow using the configured registered app.</summary>
    public Task<AccountInfo> LoginCopilotAsync(Action<global::IntelligenceX.Authentication.GitHub.GitHubDeviceAuthorization> onCode,
        CancellationToken cancellationToken = default) =>
        WithClientAsync(client => client.LoginCopilotAsync(onCode, cancellationToken), cancellationToken);

    /// <summary>Signs out of this connection's selected account.</summary>
    public Task LogoutAsync(CancellationToken cancellationToken = default) =>
        WithClientAsync(async client => { await client.LogoutAsync(cancellationToken).ConfigureAwait(false); return true; }, cancellationToken);

    private async Task<T> WithClientAsync<T>(Func<IntelligenceXClient, Task<T>> operation, CancellationToken token) {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _gate.WaitAsync(token).ConfigureAwait(false);
        try {
            ObjectDisposedException.ThrowIf(_disposed, this);
            return await operation(_client!).ConfigureAwait(false);
        } finally { _gate.Release(); }
    }

    /// <inheritdoc />
    public async Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (request.Images.Count > 0 && !Profile.SupportsImages) throw new NotSupportedException("Profile does not support images.");
            if (request.Images.Sum(image => (long)image.ByteLength) > Profile.MaxImageBytes) throw new ArgumentException("Image payload limit exceeded.", nameof(request));
            TreatmentRequest treatment = BuildTreatment(request, copyImages: true);
            if (MeasureRequestCharacters(request) > Profile.MaxRequestCharacters)
                throw new ArgumentException("Serialized treatment exceeds the profile context budget.", nameof(request));
            TreatmentResult result = await _provider.RunAsync(treatment, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (result.Text?.Length > request.MaxResponseCharacters) throw new InvalidDataException("Provider response exceeds the document output limit.");
            TurnInfo? turn = result.Raw as TurnInfo;
            return new OfficeAiExecutionResponse(result.Text ?? string.Empty, result.Id,
                turn?.Usage?.InputTokens, turn?.Usage?.OutputTokens, string.Equals(result.Status, "completed", StringComparison.OrdinalIgnoreCase));
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            throw;
        } catch (Exception exception) when (exception is not (OutOfMemoryException or InvalidDataException)) {
            throw new OfficeAiExecutionException(IntelligenceXOfficeAiErrors.Classify(exception));
        } finally { _gate.Release(); }
    }

    /// <inheritdoc />
    public int MeasureRequestCharacters(OfficeAiExecutionRequest request) => checked(request.Instructions.Length
        + TreatmentPromptBuilder.Build(BuildTreatment(request, copyImages: false)).Length
        + (Profile.EnforcesJsonSchema ? request.OutputSchema.Length : 0));

    private TreatmentRequest BuildTreatment(OfficeAiExecutionRequest request, bool copyImages) {
            var inputs = new List<TreatmentInputArtifact> {
                new() { Id = "document-request", MediaType = "application/json", Text = request.InputJson }
            };
            foreach (OfficeAiImage image in request.Images) inputs.Add(new TreatmentInputArtifact {
                Id = image.Id, Role = "source-image", MediaType = image.MediaType, ImageBytes = copyImages ? image.CopyBytes() : null
            });
            return new TreatmentRequest {
                Id = request.RequestId, Instructions = request.Instructions,
                Prompt = "Perform the operation specified in document-request using only its supplied evidence and attached images.",
                Model = Profile.Model, NewThread = true, AllowNetwork = false, InlineLocalInputFiles = false,
                Ephemeral = true,
                MaxInlineImageBytes = Profile.MaxImageBytes, Inputs = inputs.AsReadOnly(), EnforceOutputSchema = Profile.EnforcesJsonSchema,
                MaxResponseBytes = Math.Max(1_048_576L, (long)request.MaxResponseCharacters * 64),
                OutputSchema = new TreatmentOutputSchema { Contract = "officeimo.ai.response.v1", Strict = true,
                    JsonSchema = JsonLite.Parse(request.OutputSchema)!.AsObject()! }
            };
    }

    /// <summary>Disposes the owned SDK connection. Finish or cancel active operations before disposing.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _client?.Dispose();
    }
}
