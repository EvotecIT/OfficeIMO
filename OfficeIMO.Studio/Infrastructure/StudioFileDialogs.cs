namespace OfficeIMO.Studio.Infrastructure;

/// <summary>A file type offered in open and save pickers.</summary>
internal sealed record StudioFileType(string Name, IReadOnlyList<string> Extensions, string? MimeType = null) {
    public static StudioFileType Any(string name) => new(name, ["*"]);
}

/// <summary>
/// Generic open, save and folder pickers for document features. Implementations register the chosen
/// locations with Studio storage and confirm provider-backed writes before returning them.
/// </summary>
internal interface IStudioFileDialogs {
    Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken);
    Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken);
    Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken);
}

internal sealed class NoStudioFileDialogs : IStudioFileDialogs {
    public static readonly NoStudioFileDialogs Instance = new();
    public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
    public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
    public Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
}
