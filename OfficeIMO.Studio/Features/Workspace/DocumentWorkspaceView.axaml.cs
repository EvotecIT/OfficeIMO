using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView : UserControl {
    public DocumentWorkspaceView() => InitializeComponent();

    internal ListBox PagesListControl => PagesList;

    internal ListBox OrganizerListControl => OrganizerList;

    internal Button FitWidthButtonControl => FitWidthButton;

    internal Button FitPageButtonControl => FitPageButton;
}
