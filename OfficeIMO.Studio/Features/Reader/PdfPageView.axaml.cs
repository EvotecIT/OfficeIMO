using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageView : UserControl {
    private PdfPageViewModel? _viewModel;
    private bool _attached;

    public PdfPageView() {
        InitializeComponent();
        PageCanvas.LinkActivated += OnLinkActivated;
        PageCanvas.EditorGestureCompleted += OnEditorGestureCompleted;
        PageCanvas.ObjectSelected += OnObjectSelected;
        PageCanvas.TextSelectionCompleted += OnTextSelectionCompleted;
        PageCanvas.ObjectTransformCompleted += gesture => _viewModel?.TransformObject(gesture);
        PageCanvas.SizeChanged += (_, args) => _viewModel?.UpdateCanvasSize(args.NewSize);
        DataContextChanged += OnDataContextChanged;
        AttachedToVisualTree += (_, _) => {
            _attached = true;
            UpdateViewModel();
        };
        DetachedFromVisualTree += (_, _) => {
            _attached = false;
            _viewModel?.DetachFromViewport();
        };
    }

    private void OnLinkActivated(string target) => _viewModel?.ActivateLink(target);

    private void OnEditorGestureCompleted(PdfEditorGesture gesture) => _viewModel?.CompleteEditorGesture(gesture);

    private void OnObjectSelected(PdfEditorSelection? selection) => _viewModel?.SelectObject(selection);

    private Point? _contextPoint;

    private Flyout SelectionActions => (Flyout)Resources["SelectionActions"]!;

    private MenuFlyout PageMenu => (MenuFlyout)Resources["PageMenu"]!;

    // A finished text selection gets a small action bar at the pointer: copy or mark it up in one click.
    private void OnTextSelectionCompleted() {
        if (_viewModel is null || !PageCanvas.HasTextSelection) return;
        SelectionActions.ShowAt(PageCanvas, showAtPointer: true);
    }

    private void OnCanvasContextRequested(object? sender, ContextRequestedEventArgs e) {
        if (_viewModel is null) return;
        _contextPoint = e.TryGetPosition(PageCanvas, out Point point) ? point : null;
        bool hasSelection = PageCanvas.HasTextSelection;
        foreach (MenuItem item in PageMenu.Items.OfType<MenuItem>()) {
            if (Equals(item.Tag, "selection")) item.IsEnabled = hasSelection;
        }
        SelectionActions.Hide();
        PageMenu.ShowAt(PageCanvas, showAtPointer: true);
        e.Handled = true;
    }

    private async void OnCopySelectionClick(object? sender, RoutedEventArgs e) {
        SelectionActions.Hide();
        if (PageCanvas.HasTextSelection) await PageCanvas.CopySelectedTextAsync();
    }

    private void OnSelectAllClick(object? sender, RoutedEventArgs e) => PageCanvas.SelectAllPageText();

    private void OnHighlightSelectionClick(object? sender, RoutedEventArgs e) => RequestSelectionMarkup(PdfEditorTool.Highlight);

    private void OnUnderlineSelectionClick(object? sender, RoutedEventArgs e) => RequestSelectionMarkup(PdfEditorTool.Underline);

    private void OnStrikeSelectionClick(object? sender, RoutedEventArgs e) => RequestSelectionMarkup(PdfEditorTool.StrikeOut);

    private void RequestSelectionMarkup(PdfEditorTool tool) {
        SelectionActions.Hide();
        if (PageCanvas.CreateSelectionGesture() is { } gesture) _viewModel?.RequestMarkup(tool, gesture);
    }

    private void OnAddNoteClick(object? sender, RoutedEventArgs e) {
        Point point = _contextPoint ?? new Point(PageCanvas.Bounds.Width / 2D, PageCanvas.Bounds.Height / 2D);
        if (PageCanvas.CreatePointGesture(point) is { } gesture) _viewModel?.RequestMarkup(PdfEditorTool.Note, gesture);
    }

    public static readonly Avalonia.Data.Converters.IValueConverter WrapWhenMultiline =
        new Avalonia.Data.Converters.FuncValueConverter<bool, Avalonia.Media.TextWrapping>(multiline =>
            multiline ? Avalonia.Media.TextWrapping.Wrap : Avalonia.Media.TextWrapping.NoWrap);

    // The on-page editor takes focus when a page click or Tab put it there, not when the list picked the field.
    private void OnViewModelPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(PdfPageViewModel.HasInlineFormEditor) &&
            e.PropertyName != nameof(PdfPageViewModel.InlineFormField) &&
            e.PropertyName != nameof(PdfPageViewModel.FocusInlineFormEditorRequested)) return;
        FocusPendingInlineFormEditor();
    }

    private void FocusPendingInlineFormEditor() {
        if (!_attached) return;
        if (_viewModel is not { HasInlineFormEditor: true, FocusInlineFormEditorRequested: true } model) return;
        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
            if (!_attached || !ReferenceEquals(_viewModel, model) || !model.HasInlineFormEditor ||
                !model.FocusInlineFormEditorRequested) return;
            Control? editor = InlineFormText.IsVisible ? InlineFormText : InlineFormEditableChoice.IsVisible ? InlineFormEditableChoice : InlineFormCheck.IsVisible ? InlineFormCheck : InlineFormChoice.IsVisible ? InlineFormChoice : null;
            if (editor is null || !editor.Focus(NavigationMethod.Tab)) return;
            model.FocusInlineFormEditorRequested = false;
            if (editor is TextBox text) text.SelectAll();
        }, Avalonia.Threading.DispatcherPriority.Loaded);
    }

    private void OnInlineFormKeyDown(object? sender, KeyEventArgs e) {
        if (_viewModel is null || e.Key != Key.Tab) return;
        e.Handled = true;
        _viewModel.RequestInlineFormNavigation(e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1);
    }

    private void OnDataContextChanged(object? sender, EventArgs e) {
        UpdateViewModel();
    }

    private void UpdateViewModel() {
        if (ReferenceEquals(_viewModel, DataContext)) {
            if (_attached) {
                _viewModel?.AttachToViewport();
                FocusPendingInlineFormEditor();
            }
            return;
        }

        _viewModel?.DetachFromViewport();
        if (_viewModel is not null) _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        _viewModel = DataContext as PdfPageViewModel;
        if (_viewModel is not null) _viewModel.PropertyChanged += OnViewModelPropertyChanged;
        _viewModel?.UpdateCanvasSize(PageCanvas.Bounds.Size);
        if (_attached) {
            _viewModel?.AttachToViewport();
            FocusPendingInlineFormEditor();
        }
    }
}
