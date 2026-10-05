using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using System.ComponentModel;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfInlineFormWidgetView : UserControl {
    private PdfFormFieldViewModel? _listField;
    private bool _attached;
    private bool _syncingSelection;

    public PdfInlineFormWidgetView() {
        InitializeComponent();
        AddHandler(KeyDownEvent, OnEditorKeyDown, RoutingStrategies.Tunnel);
        GotFocus += (_, _) => {
            if (DataContext is PdfInlineFormWidgetViewModel widget) widget.Activate();
        };
        AttachedToVisualTree += (_, _) => { _attached = true; UpdateListSubscription(); };
        DetachedFromVisualTree += (_, _) => { _attached = false; UpdateListSubscription(); };
        DataContextChanged += (_, _) => UpdateListSubscription();
        InlineFormList.SelectionChanged += OnListSelectionChanged;
        InlineFormList.PropertyChanged += (_, change) => {
            if (change.Property == ItemsControl.ItemsSourceProperty) SyncListSelection();
        };
    }

    private void OnDropDownSelectionChanged(object? sender, SelectionChangedEventArgs args) {
        // Replacing a field's choices clears the control temporarily; that is not a user edit.
        if (DataContext is PdfInlineFormWidgetViewModel { IsDropDown: true } widget && sender is ComboBox { SelectedItem: PdfFormChoiceViewModel choice } &&
            widget.Field.Choices.Contains(choice)) widget.Field.SelectedChoice = choice;
    }

    private void UpdateListSubscription() {
        if (_listField is not null) _listField.PropertyChanged -= OnListFieldChanged;
        _listField = _attached && DataContext is PdfInlineFormWidgetViewModel { Field.IsListBoxEditor: true } widget ? widget.Field : null;
        if (_listField is not null) _listField.PropertyChanged += OnListFieldChanged;
        SyncListSelection();
    }

    private void OnListFieldChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(PdfFormFieldViewModel.DraftText)) SyncListSelection();
    }

    private bool HasCurrentList => _listField is not null && DataContext is PdfInlineFormWidgetViewModel widget &&
        ReferenceEquals(_listField, widget.Field) && ReferenceEquals(InlineFormList.ItemsSource, _listField.Choices);

    private void SyncListSelection() {
        if (_syncingSelection || !HasCurrentList) return;
        _syncingSelection = true;
        try {
            if (!_listField!.IsMultipleChoiceEditor) {
                InlineFormList.SelectedItem = _listField.SelectedChoice;
                return;
            }
            var choices = _listField.Choices.Where(choice => choice.IsSelected).ToArray();
            var selected = InlineFormList.SelectedItems!;
            if (selected.Count == choices.Length && choices.All(selected.Contains)) return;
            selected.Clear();
            foreach (var choice in choices) selected.Add(choice);
        } finally { _syncingSelection = false; }
    }

    private void OnListSelectionChanged(object? sender, SelectionChangedEventArgs args) {
        if (_syncingSelection || !HasCurrentList) return;
        _syncingSelection = true;
        try {
            if (!_listField!.IsMultipleChoiceEditor) {
                _listField.SelectedChoice = InlineFormList.SelectedItem as PdfFormChoiceViewModel;
                return;
            }
            // Selection includes virtualized rows. Container bindings would silently omit them when saving.
            var selected = InlineFormList.SelectedItems!.OfType<PdfFormChoiceViewModel>().ToHashSet();
            foreach (var choice in _listField.Choices) choice.IsSelected = selected.Contains(choice);
        } finally { _syncingSelection = false; }
    }

    internal bool FocusEditor() {
        Control? editor = new Control[] { InlineFormText, InlineFormEditableChoice, InlineFormCheck, InlineFormRadio, InlineFormChoice, InlineFormList }
            .FirstOrDefault(control => control.IsVisible);
        if (editor is null || !editor.Focus(NavigationMethod.Tab)) return false;
        if (editor is TextBox text) text.SelectAll();
        return true;
    }

    private void OnEditorKeyDown(object? sender, KeyEventArgs e) {
        if (DataContext is not PdfInlineFormWidgetViewModel widget) return;
        if (e.Key == Key.Tab) {
            e.Handled = true;
            widget.Page.RequestInlineFormNavigation(e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? -1 : 1);
        } else if (widget.Field.IsRadioButtonEditor && e.Key is Key.Left or Key.Up or Key.Right or Key.Down) {
            var radios = widget.Page.InlineFormWidgets.Where(candidate => ReferenceEquals(candidate.Field, widget.Field) && candidate.RadioChoice is not null).ToArray();
            int index = Array.IndexOf(radios, widget);
            if (index < 0) return;
            var next = radios[(index + (e.Key is Key.Left or Key.Up ? -1 : 1) + radios.Length) % radios.Length];
            next.Field.SelectedChoice = next.RadioChoice;
            var view = this.FindAncestorOfType<PdfPageView>()?.GetVisualDescendants().OfType<PdfInlineFormWidgetView>()
                .FirstOrDefault(view => ReferenceEquals(view.DataContext, next));
            view?.FocusEditor();
            e.Handled = true;
        }
    }
}
