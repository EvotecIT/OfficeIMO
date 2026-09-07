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

    private void UpdateFormAnchor() {
        foreach (var page in Pages) page.FormAnchorFieldName = IsFormsDocumentMode &&
            page == SelectedPage && SelectedFormField?.PageNumbers.Contains(page.PageNumber) == true
                ? SelectedFormField.Name : null;
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
        _refreshingFormFields = true;
        try {
            NavigateToPage(selection.PageNumber);
            SelectedFormField = field;
        } finally { _refreshingFormFields = false; ResetFormDefinition(); UpdateFormAnchor(); }
    }
}
