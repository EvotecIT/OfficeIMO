using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>A discoverable presentation command that preserves its underlying operation's guard.</summary>
public sealed class StudioCommandItem : ObservableObject, ICommand {
    private readonly ICommand _operation;
    private readonly Func<string?> _unavailable;
    private readonly Action? _prepare;

    internal StudioCommandItem(string id, string title, string description, string category,
        string shortcut, bool isTool, ICommand operation, Func<string?> unavailable, Action? prepare = null) {
        Id = id;
        Title = title;
        Description = description;
        Category = category;
        Shortcut = shortcut;
        IsTool = isTool;
        _operation = operation;
        _unavailable = unavailable;
        _prepare = prepare;
    }

    public string Id { get; }
    public string Title { get; }
    public string Description { get; }
    public string Category { get; }
    public string Shortcut { get; }
    public bool IsTool { get; }
    public string? UnavailableReason => _unavailable();
    public bool IsAvailable => UnavailableReason is null && _operation.CanExecute(null);
    public bool HasUnavailableReason => !string.IsNullOrEmpty(UnavailableReason);
    public override string ToString() => Title;
    public event EventHandler? CanExecuteChanged;

    public bool CanExecute(object? parameter) => IsAvailable;

    public async void Execute(object? parameter) => await ExecuteAsync();

    /// <summary>Rechecks availability at invocation; preparation never bypasses the operation guard.</summary>
    public async Task ExecuteAsync() {
        if (!CanExecute(null)) return;
        _prepare?.Invoke();
        if (_operation is IAsyncRelayCommand asyncCommand) await asyncCommand.ExecuteAsync(null);
        else _operation.Execute(null);
    }

    internal void Refresh() {
        OnPropertyChanged(nameof(UnavailableReason));
        OnPropertyChanged(nameof(HasUnavailableReason));
        OnPropertyChanged(nameof(IsAvailable));
        CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
