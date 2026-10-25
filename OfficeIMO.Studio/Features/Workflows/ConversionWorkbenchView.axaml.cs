using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchView : UserControl {
    private ConversionWorkbenchViewModel? _workbench;
    private bool _phoneLayout;

    private void OnOpenSavedFolderJobs(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        if (_phoneLayout) CompactTabs.SelectedItem = SetupTab;
        BatchExportPanel.IsExpanded = true;
        BatchExportPanel.BringIntoView();
        BatchExportPanel.Focus();
    }

    public ConversionWorkbenchView() {
        InitializeComponent();
        SizeChanged += OnSizeChanged;
        DataContextChanged += (_, _) => {
            if (_workbench is not null) _workbench.PropertyChanged -= OnWorkbenchChanged;
            _workbench = (DataContext as Shell.MainWindowViewModel)?.ConversionWorkbench;
            if (_workbench is not null) _workbench.PropertyChanged += OnWorkbenchChanged;
            UpdateDetailsVisibility();
        };
    }

    private void OnWorkbenchChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) {
        if (e.PropertyName is not nameof(ConversionWorkbenchViewModel.HasJobs) and not nameof(ConversionWorkbenchViewModel.IsFolderJobReviewVisible)) return;
        if (Bounds.Width > 0) ApplyResponsiveLayout(Bounds.Width);
        else UpdateDetailsVisibility();
    }

    // Details only appear once there is a job to describe, so an empty queue gets the full width.
    private void UpdateDetailsVisibility() => DetailsPanel.IsVisible = _workbench?.IsFolderJobReviewVisible != true && (_phoneLayout || _workbench?.HasJobs == true);

    internal bool IsCompactLayout { get; private set; }

    private void OnSizeChanged(object? sender, SizeChangedEventArgs e) => ApplyResponsiveLayout(e.NewSize.Width);

    internal void ApplyResponsiveLayout(double width) {
        bool phone = width < 760;
        if (phone != _phoneLayout) {
            _phoneLayout = phone;
            if (phone) {
                WorkspaceGrid.Children.Remove(SetupPanel);
                WorkspaceGrid.Children.Remove(MainReviewPanel);
                WorkspaceGrid.Children.Remove(DetailsPanel);
                SetupTab.Content = SetupPanel;
                QueueTab.Content = MainReviewPanel;
                DetailsTab.Content = DetailsPanel;
                CompactTabs.SelectedItem = _workbench?.HasJobs == true ? QueueTab : SetupTab;
            } else {
                SetupTab.Content = QueueTab.Content = DetailsTab.Content = null;
                WorkspaceGrid.Children.Add(SetupPanel);
                WorkspaceGrid.Children.Add(MainReviewPanel);
                WorkspaceGrid.Children.Add(DetailsPanel);
            }
        }
        CompactTabs.IsVisible = phone;
        IsCompactLayout = width < 1100D;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(phone ? 0 : IsCompactLayout ? 238D : Math.Clamp(width * .2D, 270D, 360D));
        bool showDetailsColumn = !IsCompactLayout && _workbench?.HasJobs == true && !_workbench.IsFolderJobReviewVisible;
        WorkspaceGrid.ColumnDefinitions[2].Width = new GridLength(showDetailsColumn ? Math.Clamp(width * .25D, 320D, 480D) : 0D);
        Grid.SetRow(DetailsPanel, IsCompactLayout ? 1 : 0);
        Grid.SetColumn(DetailsPanel, IsCompactLayout ? 1 : 2);
        Grid.SetColumnSpan(DetailsPanel, 1);
        DetailsPanel.MaxHeight = IsCompactLayout && !phone ? 230D : double.PositiveInfinity;
        UpdateDetailsVisibility();
    }
}
