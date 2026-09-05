using Avalonia.Controls;
using Avalonia.Input;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView : UserControl {
    public DocumentWorkspaceView() => InitializeComponent();

    internal ListBox PagesListControl => PagesList;

    internal ListBox OrganizerListControl => OrganizerList;

    internal ListBox GridPagesListControl => GridPagesList;

    internal Button FitWidthButtonControl => FitWidthButton;

    internal Button FitPageButtonControl => FitPageButton;

    internal void FocusSearch() {
        NavigationTabs.SelectedIndex = 2;
        SearchBox.Focus();
        SearchBox.SelectAll();
    }

    private void OnGridPagePointerPressed(object? sender, PointerPressedEventArgs e) {
        if (!e.GetCurrentPoint(GridPagesList).Properties.IsLeftButtonPressed ||
            DataContext is not MainWindowViewModel viewModel) return;

        Control? control = e.Source as Control;
        while (control is not null) {
            if (control.DataContext is PdfPageViewModel page) {
                viewModel.SelectedPage = page;
                return;
            }
            control = control.Parent as Control;
        }
    }
}
