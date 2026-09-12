namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Non-secret setup choices. Credentials belong to the provider's account store or session memory.</summary>
internal sealed record StudioAiConnectionPreference {
    public string Provider { get; init; } = "chatgpt";
    public string Model { get; init; } = string.Empty;
    public string? Endpoint { get; init; }
    public string? GitHubClientId { get; init; }
}
