using Avalonia.Controls;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class OutputIntakeWorkbenchView : UserControl {
    public OutputIntakeWorkbenchView() {
        InitializeComponent();
        Adapt(PrintLayout, "Conversion.CompactSetup", "Conversion.CompactDetails");
        Adapt(ExportLayout, "Conversion.CompactSetup", "Conversion.CompactDetails");
        Adapt(AssemblyLayout, "Conversion.CompactQueue", "Conversion.CompactSetup");
    }

    private void Adapt(Grid layout, string firstTitle, string secondTitle) {
        Control first = layout.Children[0], second = layout.Children[1];
        var firstTab = new TabItem { Header = StudioLocalization.Current.Get(firstTitle) };
        var secondTab = new TabItem { Header = StudioLocalization.Current.Get(secondTitle) };
        var tabs = new TabControl { Classes = { "pageTabs" }, Items = { firstTab, secondTab } };
        Grid.SetColumnSpan(tabs, 2);
        bool compact = false;
        SizeChanged += (_, e) => {
            bool next = e.NewSize.Width < 760;
            if (compact == next) return;
            compact = next;
            if (compact) {
                layout.Children.Clear();
                firstTab.Content = first;
                secondTab.Content = second;
                layout.Children.Add(tabs);
            } else {
                layout.Children.Clear();
                firstTab.Content = secondTab.Content = null;
                layout.Children.Add(first);
                layout.Children.Add(second);
            }
        };
    }
}
