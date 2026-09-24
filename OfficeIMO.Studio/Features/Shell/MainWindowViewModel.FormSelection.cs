using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private bool _refreshingFormFields;

    private void RebuildFormFields() {
        var previous = SelectedFormField;
        bool sameDocument = ReferenceEquals(_formWorkspace, _workspace);
        _refreshingFormFields = true;
        try { RebuildFormFieldModels(); }
        finally {
            _refreshingFormFields = false;
            if (!sameDocument || previous is null || SelectedFormField?.HasSameDefinition(previous) != true) ResetFormDefinition();
            UpdateFormAnchor();
        }
    }

    private bool _focusInlineFormEditor;

    private void UpdateFormAnchor() {
        foreach (var page in Pages) {
            bool anchored = IsFormsDocumentMode && page == SelectedPage && SelectedFormField?.PageNumbers.Contains(page.PageNumber) == true;
            page.FormAnchorFieldName = anchored ? SelectedFormField!.Name : null;
            page.ShowInlineFormField(anchored && CanFillInPlace(SelectedFormField!) ? SelectedFormField : null, _focusInlineFormEditor);
        }
        _focusInlineFormEditor = false;
    }

    // Text, check box and drop-down fields with one widget are filled on the page; others stay in the side pane.
    private bool CanFillInPlace(PdfFormFieldViewModel field) => _workspace?.CanFillForms == true && field.CanFill && !IsFillSignPlacement(ActiveEditorTool) &&
        (field.IsTextEditor || field.SingleWidget is not null && (field.IsCheckBoxEditor || field.IsSingleChoiceEditor));

    // Tab order follows the document's field order, which Studio already keeps in tab order per page.
    private void OnInlineFormNavigationRequested(int direction) {
        PdfFormFieldViewModel[] fillable = FormFields.Where(field => field.PageNumbers.Count > 0 && CanFillInPlace(field)).ToArray();
        if (fillable.Length == 0) return;
        int index = SelectedFormField is null ? -1 : Array.IndexOf(fillable, SelectedFormField);
        int next = index < 0 ? (direction > 0 ? 0 : fillable.Length - 1) : (index + direction + fillable.Length) % fillable.Length;
        PdfFormFieldViewModel field = fillable[next];
        _focusInlineFormEditor = true;
        _refreshingFormFields = true;
        try {
            if (SelectedPage?.PageNumber is not int page || !field.PageNumbers.Contains(page)) NavigateToPage(field.PageNumbers[0]);
            SelectedFormField = field;
        } finally { _refreshingFormFields = false; ResetFormDefinition(); UpdateFormAnchor(); }
    }

    [RelayCommand]
    private void ShowSelectedFormField() {
        if (SelectedFormField?.PageNumbers.FirstOrDefault() is not > 0) return;
        NavigateToPage(SelectedFormField.PageNumbers[0]);
        if (SelectedPage is not null) SelectedPage.FormAnchorFieldName = null;
        UpdateFormAnchor();
    }

    private void SelectFormWidget(PdfEditorSelection selection) {
        var field = FormFields.FirstOrDefault(field => field.Name == selection.FieldName);
        if (field is null) return;
        ClearObjectSelection();
        DocumentMode = StudioDocumentMode.Forms;
        _focusInlineFormEditor = true;
        _refreshingFormFields = true;
        try {
            NavigateToPage(selection.PageNumber);
            SelectedFormField = field;
        } finally { _refreshingFormFields = false; ResetFormDefinition(); UpdateFormAnchor(); }
    }
}
