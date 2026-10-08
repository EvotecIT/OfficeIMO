using Avalonia;
using Avalonia.Controls;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Desktop window presentation for the same review content used by mobile sheets.</summary>
public class StudioDialogWindow : Window {
    protected StudioDialogWindow(StudioDialogContent view) {
        Width = view.Width;
        Height = view.Height;
        MinWidth = view.MinWidth;
        MinHeight = view.MinHeight;
        CanResize = view.CanResize;
        SizeToContent = view.SizeToContent;
        WindowStartupLocation = view.WindowStartupLocation;
        Background = view.Background;
        FontSize = view.FontSize;
        DataContext = view.DataContext;
        this.Bind(TitleProperty, view.GetObservable(StudioDialogContent.TitleProperty));
        view.ClearValue(WidthProperty);
        view.ClearValue(HeightProperty);
        view.ClearValue(MinWidthProperty);
        view.ClearValue(MinHeightProperty);
        Content = view;
        NameScope.SetNameScope(this, NameScope.GetNameScope(view));
        view.CloseRequested += Close;
        Opened += (_, _) => view.Presented(Owner ?? this);
        Closed += (_, _) => { view.CloseRequested -= Close; view.Dismissed(); };
    }
}
