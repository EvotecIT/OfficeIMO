using System.Globalization;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Touch page entry presented by the existing keyboard-aware mobile dialog host.</summary>
public sealed partial class PageNavigationDialogContent : StudioDialogContent {
    private readonly int _pageCount;

    public PageNavigationDialogContent() { InitializeComponent(); }

    internal PageNavigationDialogContent(int currentPage, int pageCount) : this() {
        _pageCount = pageCount;
        PageRange.Text = StudioLocalization.Current.Format("Reader.PageRange", pageCount);
        ValidationMessage.Text = StudioLocalization.Current.Format("Reader.InvalidPage", pageCount);
        PageInput.Text = currentPage.ToString(CultureInfo.CurrentCulture);
        Opened += (_, _) => { PageInput.Focus(); PageInput.SelectAll(); };
    }

    private void OnPageTextChanged(object? sender, TextChangedEventArgs e) {
        if (GoButton is null) return;
        bool valid = PdfPageNavigation.TryParsePageNumber(PageInput.Text, _pageCount, out _);
        GoButton.IsEnabled = valid;
        ValidationMessage.IsVisible = !valid && !string.IsNullOrWhiteSpace(PageInput.Text);
    }

    private void OnPageKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key != Key.Enter) return;
        Submit();
        e.Handled = true;
    }

    private void OnGo(object? sender, RoutedEventArgs e) => Submit();
    private void OnCancel(object? sender, RoutedEventArgs e) => Close();

    private void Submit() {
        if (PdfPageNavigation.TryParsePageNumber(PageInput.Text, _pageCount, out int page)) Close(page);
    }
}
