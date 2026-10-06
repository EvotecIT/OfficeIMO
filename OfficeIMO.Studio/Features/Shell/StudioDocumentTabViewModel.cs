using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Adapts shared document-tab state to Avalonia commands and localized labels.</summary>
public sealed partial class StudioDocumentTabViewModel : ObservableObject, IStudioDocumentTab<MainWindowViewModel> {
    private readonly StudioDocumentTabState<MainWindowViewModel> _state;
    private readonly Func<StudioDocumentTabViewModel, Task> _close;
    internal StudioDocumentTabViewModel(MainWindowViewModel document, Func<StudioDocumentTabViewModel, Task> close) {
        _state = new(document);
        _close = close;
        _state.PropertyChanged += (_, args) => {
            OnPropertyChanged(args.PropertyName);
            if (args.PropertyName == nameof(DisplayTitle)) OnPropertyChanged(nameof(CloseLabel));
        };
    }
    internal MainWindowViewModel Document => _state.Document;
    MainWindowViewModel IStudioDocumentTab<MainWindowViewModel>.Document => Document;
    public string Title { get => _state.Title; set => _state.Title = value; }
    public string DisplayTitle => _state.DisplayTitle;
    public string CloseLabel => Infrastructure.Localization.StudioLocalization.Current.Format("Tabs.CloseDocument", DisplayTitle);
    public bool IsDirty => _state.IsDirty;
    public string? SourcePath => _state.SourcePath;
    public override string ToString() => Title;
    [RelayCommand] private Task CloseAsync() => _close(this);
    public void Dispose() => _state.Dispose();
}
