using Avalonia;
using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Settings;

public sealed partial class SettingsView : UserControl {
    private const double CompactWidth = 720;
    private bool _followingScroll;
    private double? _requestedOffset;

    public SettingsView() {
        InitializeComponent();
        FocusReadingHelp.IsVisible = !OperatingSystem.IsIOS();
        SectionList.SelectedIndex = 0;
        SectionList.SelectionChanged += (_, _) => ScrollToSelectedSection();
        SectionScroller.ScrollChanged += (_, _) => FollowScrolledSection();
        SizeChanged += (_, e) => UseCompactLayout(e.NewSize.Width < CompactWidth);
    }

    internal void UseTouchPresentation() => FocusReadingHelp.IsVisible = false;

    private IEnumerable<Control> Sections =>
        SectionList.Items.OfType<ListBoxItem>().Select(item => this.FindControl<Control>((string)item.Tag!)!);

    private void UseCompactLayout(bool compact) {
        SectionNavigation.IsVisible = !compact;
        SettingsLayout.ColumnDefinitions = new ColumnDefinitions(compact ? "0,*" : "220,*");
        SettingsContent.Classes.Set("compact", compact);
    }

    private void ScrollToSelectedSection() {
        if (SectionList.SelectedItem is not ListBoxItem { Tag: string name } || this.FindControl<Control>(name) is not { } section) return;
        if (_followingScroll || section.TranslatePoint(default, SettingsContent) is not { } top) return;
        _requestedOffset = Math.Max(0, top.Y + SettingsContent.Margin.Top - 16);
        SectionScroller.Offset = new Vector(0, _requestedOffset.Value);
    }

    // Keep the list on the section being read, so the highlighted entry always matches the page.
    private void FollowScrolledSection() {
        if (!SectionNavigation.IsVisible) return;
        double maximum = Math.Max(0, SectionScroller.Extent.Height - SectionScroller.Viewport.Height);
        // A chosen section near the end cannot reach the top; keep it selected until the reader scrolls away.
        if (_requestedOffset is { } requested && Math.Abs(SectionScroller.Offset.Y - Math.Min(requested, maximum)) < 1) return;
        _requestedOffset = null;
        // A section counts as being read once its heading reaches the upper third of the page.
        double reading = SectionScroller.Offset.Y + Math.Min(SectionScroller.Viewport.Height / 3, 200);
        bool atEnd = SectionScroller.Offset.Y >= maximum - 1;
        int current = 0, index = 0;
        foreach (Control section in Sections) {
            if (section.TranslatePoint(default, SettingsContent) is { } top && top.Y + SettingsContent.Margin.Top <= reading) current = index;
            index++;
        }
        if (atEnd && SectionScroller.Offset.Y > 0) current = index - 1;
        if (SectionList.SelectedIndex == current) return;
        _followingScroll = true;
        try {
            SectionList.SelectedIndex = current;
        } finally {
            _followingScroll = false;
        }
    }
}
