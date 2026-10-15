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
        string shortcut, bool isTool, ICommand operation, Func<string?> unavailable, Action? prepare = null, string? categoryId = null,
        IReadOnlyList<Avalonia.Input.KeyGesture>? gestures = null, bool runsInTextEntry = false) {
        Id = id;
        CategoryId = categoryId ?? category;
        Gestures = gestures ?? [];
        RunsInTextEntry = runsInTextEntry;
        // Real gestures describe themselves; a context hint such as "Alt+↑ (page list)" stays as written.
        if (Gestures.Count > 0) shortcut = StudioShortcuts.Describe(Gestures[0]);
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
    /// <summary>The untranslated category used to group and order commands.</summary>
    public string CategoryId { get; }
    public string Shortcut { get; }
    /// <summary>Key gestures that run this command from anywhere in the window; the first is the one shown.</summary>
    public IReadOnlyList<Avalonia.Input.KeyGesture> Gestures { get; }
    /// <summary>True when the gesture still runs while a text field has focus (file commands), false when the field owns it (undo, zoom).</summary>
    public bool RunsInTextEntry { get; }
    public Avalonia.Input.KeyGesture? Gesture => Gestures.Count > 0 ? Gestures[0] : null;
    /// <summary>Tooltip text for toolbar buttons: the command title with its shortcut, from the same definition the key handler uses.</summary>
    public string ToolTipText => string.IsNullOrEmpty(Shortcut) ? Title : Title + " (" + Shortcut + ")";
    /// <summary>Toolbar tooltip that also says why the command cannot run right now.</summary>
    public string StatusToolTip => UnavailableReason is { Length: > 0 } reason ? ToolTipText + Environment.NewLine + reason : ToolTipText;
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
        OnPropertyChanged(nameof(StatusToolTip));
        CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
