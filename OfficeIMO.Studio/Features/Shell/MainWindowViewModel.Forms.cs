using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty]
    private string _newFormFieldDisplayName = string.Empty;

    [ObservableProperty]
    private bool _newFormFieldIsRequired;

    [ObservableProperty]
    private bool _newFormFieldIsReadOnly;

    [ObservableProperty]
    private int _newFormFieldMaxLength;

    public ObservableCollection<PdfFormFieldViewModel> UnassignedFormDrafts { get; } = [];
    public bool HasUnassignedFormDrafts => UnassignedFormDrafts.Count > 0;
    public bool HasFormDrafts => FormFields.Any(entry => entry.HasDraft) || HasUnassignedFormDrafts;
    public bool CanApplyFormDrafts => !IsWorkspaceBusy && _workspace?.CanFillForms == true && HasFormDrafts &&
        !HasUnassignedFormDrafts && FormFields.Where(entry => entry.HasDraft).All(entry => entry.CanApplyValue);

    private void NotifyFormDraftState() {
        OnPropertyChanged(nameof(CanEditFormDefinition));
        OnPropertyChanged(nameof(CanSetFormTabOrder));
        OnPropertyChanged(nameof(CanMoveFormWidget));
        OnPropertyChanged(nameof(CanSetFormDefault));
        OnPropertyChanged(nameof(HasFormDrafts));
        OnPropertyChanged(nameof(HasUnassignedFormDrafts));
        OnPropertyChanged(nameof(CanApplyFormDrafts));
        OnPropertyChanged(nameof(IsDirty));
        if (_workspace is not null) DocumentName = _workspace.FileName + (IsDirty ? " *" : string.Empty);
    }

    private Dictionary<string, PdfFormFieldValue>? CaptureFormDrafts() {
        if (HasUnassignedFormDrafts) {
            ErrorMessage = "A field with an unapplied draft was removed or changed. Restore the field with Undo or discard its retained draft before saving.";
            return null;
        }
        var drafts = FormFields.Where(field => field.HasDraft).ToArray();
        if (drafts.Length > 0 && _workspace?.CanFillForms != true) {
            ErrorMessage = "This document does not permit applying form values.";
            return null;
        }
        var invalid = drafts.FirstOrDefault(field => !field.CanApplyValue);
        if (invalid is not null) {
            SelectedFormField = invalid;
            ErrorMessage = invalid.DisplayName + ": " + (invalid.HasDraftConflict
                ? "The saved field changed. Choose whether to keep your draft or reset it."
                : invalid.ValidationSummary);
            return null;
        }
        return drafts.ToDictionary(field => field.Name, field => field.CreateValue(), StringComparer.Ordinal);
    }

    [RelayCommand]
    private async Task ApplyFormDraftsAsync(CancellationToken token) {
        if (_workspace is null || IsWorkspaceBusy) return;
        var values = CaptureFormDrafts();
        if (values is null || values.Count == 0) return;
        var workspace = _workspace;
        await RunMutationAsync(cancellation => ApplyCapturedFormValuesAsync(values,
            () => workspace.FillFormFieldsAsync(values, cancellation, CreateProgress())), token).ConfigureAwait(true);
    }

    private async Task ApplyCapturedFormValuesAsync(IReadOnlyDictionary<string, PdfFormFieldValue> values, Func<Task> operation) {
        var fields = FormFields.Where(field => values.ContainsKey(field.Name)).ToArray();
        await operation().ConfigureAwait(true);
        // Advance the saved baseline before rebuilding, including when a newer edit returned to the old saved value.
        foreach (var field in fields) field.RecordAppliedValue(values[field.Name]);
    }

    [RelayCommand]
    private void DiscardUnassignedFormDraft(PdfFormFieldViewModel? draft) {
        if (IsWorkspaceBusy || draft is null) return;
        UnassignedFormDrafts.Remove(draft);
        NotifyFormDraftState();
    }
}
