using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using IntelligenceX.OpenAI.Auth;
using OfficeIMO.AI;
using OfficeIMO.AI.IntelligenceX;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Assistant;

/// <summary>Application connection settings. Credentials remain in IX's Studio-specific store or session memory.</summary>
internal sealed partial class StudioAiConnections : ObservableObject {
    private readonly FileAuthBundleStore _chatGptStore;
    private readonly FileAuthBundleStore _copilotStore;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _operation;
    private long _revision;

    internal StudioAiConnections(string root, IStudioLocalizer localizer) {
        _localizer = localizer;
        _chatGptStore = new FileAuthBundleStore(Path.Combine(root, "AI", "chatgpt-auth.json"));
        _copilotStore = new FileAuthBundleStore(Path.Combine(root, "AI", "copilot-auth.json"));
        Status = Text("NotConnected", "Choose a connection. No document is sent during sign-in or model discovery.");
    }

    internal event EventHandler? Changed;
    internal long Revision => _revision;
    internal Func<Uri, Task>? OpenUri { get; set; }
    public IReadOnlyList<string> Providers { get; } = ["ChatGPT", "OpenAI-compatible API", "GitHub Copilot", "Local model"];
    public ObservableCollection<string> Models { get; } = [];
    public ObservableCollection<string> Accounts { get; } = [];
    public bool HasSavedAccounts => Accounts.Count > 0;
    [ObservableProperty] private int _providerIndex;
    [ObservableProperty] private string _model = string.Empty;
    [ObservableProperty] private string _endpoint = "https://api.openai.com/v1/";
    [ObservableProperty] private string _apiKey = string.Empty;
    [ObservableProperty] private string _accountId = string.Empty;
    [ObservableProperty] private string _gitHubClientId = string.Empty;
    [ObservableProperty] private string _status = string.Empty;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _isConnected;
    public bool IsChatGpt => ProviderIndex == 0;
    public bool IsCopilot => ProviderIndex == 2;
    public bool IsLocal => ProviderIndex == 3;
    public bool UsesEndpoint => ProviderIndex is 1 or 3;
    public bool UsesCredential => ProviderIndex is 1 or 2 or 3;
    public bool CanSignIn => ProviderIndex == 0 || (ProviderIndex == 2 && !string.IsNullOrWhiteSpace(GitHubClientId));
    public bool CanConnect => !IsBusy;
    public bool CanUse => IsConnected && !IsBusy && !string.IsNullOrWhiteSpace(Model);
    public string Summary => string.IsNullOrWhiteSpace(Model) ? Providers[ProviderIndex] : $"{Providers[ProviderIndex]} · {Model}";

    partial void OnProviderIndexChanged(int value) {
        AccountId = string.Empty; ApiKey = string.Empty; Model = string.Empty; Models.Clear(); Accounts.Clear();
        OnPropertyChanged(nameof(HasSavedAccounts));
        Endpoint = value == 3 ? "http://localhost:11434/v1/" : "https://api.openai.com/v1/";
        foreach (string name in new[] { nameof(IsChatGpt), nameof(IsCopilot), nameof(IsLocal), nameof(UsesEndpoint), nameof(UsesCredential), nameof(CanSignIn) }) OnPropertyChanged(name);
        Invalidate();
    }
    partial void OnModelChanged(string value) => Invalidate(keepConnection: true);
    partial void OnEndpointChanged(string value) => Invalidate();
    partial void OnApiKeyChanged(string value) => Invalidate();
    partial void OnAccountIdChanged(string value) => Invalidate();
    partial void OnGitHubClientIdChanged(string value) { OnPropertyChanged(nameof(CanSignIn)); Invalidate(); }
    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanConnect)); OnPropertyChanged(nameof(CanUse));
        ConnectCommand.NotifyCanExecuteChanged(); SignInCommand.NotifyCanExecuteChanged(); ImportCodexLoginCommand.NotifyCanExecuteChanged();
        SignOutCommand.NotifyCanExecuteChanged();
    }
    partial void OnIsConnectedChanged(bool value) => OnPropertyChanged(nameof(CanUse));

    private void Invalidate(bool keepConnection = false) {
        _revision++;
        _operation?.Cancel();
        if (!keepConnection) IsConnected = false;
        OnPropertyChanged(nameof(CanUse)); Changed?.Invoke(this, EventArgs.Empty);
        OnPropertyChanged(nameof(Summary));
    }

    internal OfficeAiExecutionProfile Profile() => new() {
        Id = "studio-document", Provider = Providers[ProviderIndex], Model = Model.Trim(), IsLocal = IsLocal,
        SupportsImages = false, EnforcesJsonSchema = false
    };

    internal OfficeAiIntelligenceXOptions Options(bool importCodex = false) => new() {
        Transport = ProviderIndex switch { 0 => OfficeAiIntelligenceXTransport.ChatGpt, 2 => OfficeAiIntelligenceXTransport.CopilotNative, _ => OfficeAiIntelligenceXTransport.CompatibleHttp },
        Endpoint = UsesEndpoint ? new Uri(Endpoint.Trim(), UriKind.Absolute) : null,
        ApiKey = UsesEndpoint && !string.IsNullOrWhiteSpace(ApiKey) ? ApiKey : null,
        AuthStore = _chatGptStore, AccountId = importCodex || string.IsNullOrWhiteSpace(AccountId) ? null : AccountId.Trim(),
        PreferCurrentCodexSession = importCodex,
        LoadCodexAuthJson = false,
        CopilotOptions = IsCopilot ? new() {
            AuthStore = _copilotStore, AccountId = string.IsNullOrWhiteSpace(AccountId) ? null : AccountId.Trim(),
            GitHubToken = string.IsNullOrWhiteSpace(ApiKey) ? null : ApiKey,
            GitHubClientId = string.IsNullOrWhiteSpace(GitHubClientId) ? null : GitHubClientId.Trim(),
            UseEnvironmentCredentials = false, Streaming = true
        } : null,
        Streaming = true
    };

    [RelayCommand(CanExecute = nameof(CanConnect))]
    private Task ConnectAsync() => RunConnectionAsync(signIn: false, importCodex: false, signOut: false);
    [RelayCommand(CanExecute = nameof(CanConnect))]
    private Task SignInAsync() => RunConnectionAsync(signIn: true, importCodex: false, signOut: false);
    [RelayCommand(CanExecute = nameof(CanConnect))]
    private Task ImportCodexLoginAsync() => RunConnectionAsync(signIn: false, importCodex: true, signOut: false);
    [RelayCommand(CanExecute = nameof(CanConnect))]
    private Task SignOutAsync() => RunConnectionAsync(signIn: false, importCodex: false, signOut: true);
    [RelayCommand] private void Cancel() => _operation?.Cancel();

    private async Task RunConnectionAsync(bool signIn, bool importCodex, bool signOut) {
        if (IsBusy || (importCodex && !IsChatGpt) || (signIn && !CanSignIn)) return;
        Invalidate();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(5));
        _operation = cancellation; IsBusy = true;
        long revision = _revision;
        try {
            Status = Text("Connecting", "Connecting…");
            if (UsesEndpoint && (!Uri.TryCreate(Endpoint.Trim(), UriKind.Absolute, out var endpoint)
                || endpoint.Scheme is not ("http" or "https") || endpoint.UserInfo.Length > 0
                || endpoint.Query.Length > 0 || endpoint.Fragment.Length > 0
                || IsLocal && !endpoint.IsLoopback || endpoint.Scheme == "http" && !endpoint.IsLoopback)) {
                Status = Text("InvalidEndpoint", "Enter an HTTP(S) base URL without credentials or query parameters. Local models require a loopback address; remote endpoints require HTTPS.");
                return;
            }
            if (!signIn && !importCodex && !UsesEndpoint && string.IsNullOrWhiteSpace(ApiKey)) {
                var store = IsChatGpt ? _chatGptStore : _copilotStore;
                var accounts = await store.ListAsync(IsChatGpt ? "openai-codex" : "copilot", cancellation.Token);
                cancellation.Token.ThrowIfCancellationRequested();
                if (revision != _revision) return;
                RefreshAccounts(accounts.Select(account => account.AccountId));
                if (accounts.Count > 1 && string.IsNullOrWhiteSpace(AccountId)) {
                    Status = Text("ChooseAccount", "Several Studio accounts are saved. Select an account before connecting or signing out.");
                    return;
                }
                if (!signOut && accounts.Count == 0) {
                    Status = Text("NoSavedAccount", "No Studio login is saved. Use Sign in, import an existing Codex login for ChatGPT, or supply a Copilot credential.");
                    return;
                }
            }
            if (importCodex) {
                AuthBundle bundle = CodexAuthStore.TryReadBundle(CodexAuthStore.ResolveAuthPath())
                    ?? throw new InvalidOperationException("No local Codex login is available.");
                await _chatGptStore.SaveAsync(bundle, cancellation.Token);
            }
            // The SDK requires an identifier before discovery; this placeholder is never used for inference.
            OfficeAiExecutionProfile profile = Profile() with { Model = string.IsNullOrWhiteSpace(Model) ? "model-discovery" : Model.Trim() };
            using var client = await IntelligenceXOfficeAiExecutor.ConnectAsync(profile, Options(importCodex), cancellation.Token);
            if (signOut) {
                await client.LogoutAsync(cancellation.Token);
                if (revision != _revision) return;
                ApiKey = string.Empty; AccountId = string.Empty; Models.Clear();
                Status = Text("SignedOut", "Signed out of the selected Studio connection.");
                return;
            }
            if (signIn && IsChatGpt) await client.LoginChatGptAsync(url => ShowLogin(url, null, revision, cancellation.Token), cancellation.Token);
            if (signIn && IsCopilot) await client.LoginCopilotAsync(code => ShowLogin(code.VerificationUri.AbsoluteUri, code.UserCode, revision, cancellation.Token), cancellation.Token);
            var models = await client.ListModelsAsync(cancellation.Token);
            var account = IsChatGpt || IsCopilot ? await client.GetAccountAsync(cancellation.Token) : null;
            cancellation.Token.ThrowIfCancellationRequested();
            if (revision != _revision) return;
            _operation = null;
            if (account is not null) AccountId = account.AccountId ?? string.Empty;
            Models.Clear();
            foreach (var candidate in models.Models) {
                string id = candidate.Model;
                if (!string.IsNullOrWhiteSpace(id) && !Models.Contains(id)) Models.Add(id);
            }
            if (string.IsNullOrWhiteSpace(Model) && Models.Count > 0) Model = Models[0];
            IsConnected = true;
            Status = account is null ? Text("Connected", "Connected. Select or enter a model before asking a question.")
                : _localizer.FormatOrDefault("Assistant.AccountConnected", "Connected as {0} ({1}).", account.Email ?? account.AccountId ?? "account", account.PlanType ?? Providers[ProviderIndex]);
        } catch (OperationCanceledException) {
            Status = Text("Cancelled", "Connection cancelled or timed out. Reconnect to inspect saved sign-in state; no new model was selected.");
        } catch (Exception exception) {
            // Provider exceptions may contain credentials, response bodies or authorization URLs.
            Status = StudioAiFailureText.FromException(_localizer, exception);
        } finally { if (ReferenceEquals(_operation, cancellation)) _operation = null; IsBusy = false; }
    }

    internal void RefreshAccounts(IEnumerable<string?> accountIds) {
        var available = accountIds.Where(id => !string.IsNullOrWhiteSpace(id)).Select(id => id!).Distinct(StringComparer.Ordinal).ToArray();
        // Keep surviving items in place so a bound selection cannot transiently clear
        // AccountId and cancel the connection operation that refreshed this list.
        foreach (string id in available) if (!Accounts.Contains(id)) Accounts.Add(id);
        for (int index = Accounts.Count - 1; index >= 0; index--)
            if (!available.Contains(Accounts[index], StringComparer.Ordinal)) Accounts.RemoveAt(index);
        OnPropertyChanged(nameof(HasSavedAccounts));
    }

    private void ShowLogin(string url, string? code, long revision, CancellationToken token) => Dispatcher.UIThread.Post(async () => {
        if (token.IsCancellationRequested || revision != _revision) return;
        Status = code is null ? Text("BrowserSignIn", "Complete sign-in in your browser.")
            : _localizer.FormatOrDefault("Assistant.DeviceCode", "Enter code {0} in your browser.", code);
        try { if (OpenUri is not null) await OpenUri(new Uri(url)); }
        catch { Status = Text("BrowserFailed", "The sign-in browser could not be opened. Cancel and try again."); }
    });

    private string Text(string key, string fallback) => _localizer.GetOrDefault("Assistant." + key, fallback);
}
