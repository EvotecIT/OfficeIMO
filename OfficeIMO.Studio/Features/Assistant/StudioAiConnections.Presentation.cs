using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Assistant;

/// <summary>A provider's setup guidance; transport and authentication remain owned by IntelligenceX.</summary>
internal sealed record StudioAiProvider(string Key, string Name, string Category, string Description, string Preparation, string? Endpoint) {
    public override string ToString() => Name;
}

internal sealed partial class StudioAiConnections {
    private readonly StudioPreferencesService? _preferences;
    public IReadOnlyList<StudioAiProvider> ProviderChoices { get; }
    public StudioAiProvider SelectedProvider => ProviderChoices[ProviderIndex];
    public bool IsApi => ProviderIndex == 1;
    public string ConnectionHeading => _localizer.Get(IsConnected ? "Connections.Connected" : "Connections.Setup");
    [ObservableProperty] private bool _showAdvanced;

    private void RestoreSelection() {
        if (_preferences?.Current.AiConnection is not { } saved) return;
        int index = ProviderChoices.ToList().FindIndex(provider => provider.Key == saved.Provider);
        if (index < 0) return;
        ProviderIndex = index;
        Model = saved.Model ?? string.Empty;
        if (UsesEndpoint && Uri.TryCreate(saved.Endpoint, UriKind.Absolute, out var endpoint)
            && endpoint.Scheme is "https" or "http" && endpoint.UserInfo.Length == 0
            && endpoint.Query.Length == 0 && endpoint.Fragment.Length == 0
            && (!IsLocal || endpoint.IsLoopback) && (endpoint.Scheme == "https" || endpoint.IsLoopback)) Endpoint = endpoint.AbsoluteUri;
        GitHubClientId = saved.GitHubClientId ?? string.Empty;
    }

    internal bool RememberSelection() {
        if (!CanUse) return false;
        try {
            _preferences?.Update(current => current with {
                AiConnection = new StudioAiConnectionPreference {
                    Provider = SelectedProvider.Key, Model = Model.Trim(),
                    Endpoint = UsesEndpoint ? Endpoint.Trim() : null,
                    GitHubClientId = IsCopilot ? GitHubClientId.Trim() : null
                }
            });
            return true;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            Status = _localizer.Get("Connections.PreferencesFailed");
            return false;
        }
    }

    private static IReadOnlyList<StudioAiProvider> CreateProviderChoices(IStudioLocalizer localizer) {
        StudioAiProvider Provider(string name, string key, string category, string? endpoint = null) =>
            new(key.ToLowerInvariant(), name, localizer.Get("Connections." + category), localizer.Get("Connections." + key + ".Description"),
                localizer.Get("Connections." + key + ".Preparation"), endpoint);
        return [
            Provider("ChatGPT", "ChatGpt", "Account"),
            Provider("OpenAI-compatible API", "Api", "Cloud", "https://api.openai.com/v1/"),
            Provider("GitHub Copilot", "Copilot", "Account"),
            Provider(localizer.Get("Connections.CustomLocal"), "Local", "OnDevice", "http://localhost:11434/v1/"),
            Provider("LM Studio", "LmStudio", "OnDevice", "http://localhost:1234/v1/"),
            Provider("Ollama", "Ollama", "OnDevice", "http://localhost:11434/v1/")
        ];
    }
}
