using Avalonia;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private static readonly PdfEditorToolChoice[] AvailableEditorTools = {
        new(PdfEditorTool.Select, "Select", "Select and copy text or open links"),
        new(PdfEditorTool.Note, "Note", "Click to add a comment note"),
        new(PdfEditorTool.FreeText, "Text box", "Draw a free-text annotation"),
        new(PdfEditorTool.Highlight, "Highlight", "Drag across text or an area"),
        new(PdfEditorTool.Underline, "Underline", "Drag across text or an area"),
        new(PdfEditorTool.StrikeOut, "Strikeout", "Drag across text or an area"),
        new(PdfEditorTool.Rectangle, "Rectangle", "Draw a rectangle annotation"),
        new(PdfEditorTool.Ellipse, "Ellipse", "Draw an ellipse annotation"),
        new(PdfEditorTool.Line, "Line", "Drag a review line"),
        new(PdfEditorTool.Ink, "Ink", "Draw a freehand ink path"),
        new(PdfEditorTool.Stamp, "Stamp", "Place an annotation stamp"),
        new(PdfEditorTool.AddText, "Add text", "Add permanent page text without reflowing existing content"),
        new(PdfEditorTool.AddImage, "Add image", "Choose and place a PNG or JPEG image"),
        new(PdfEditorTool.Link, "Link", "Draw a URI link hotspot"),
        new(PdfEditorTool.SignatureAppearance, "Signature appearance", "Draw a visual-only signature label; this does not cryptographically sign the PDF"),
        new(PdfEditorTool.Redact, "Redact", "Draw an area, review it, then permanently remove intersecting content")
    };

    private PdfEditorGesture? _pendingRedaction;
    private PdfRedactionPlan? _pendingRedactionPlan;
    private PdfWorkspace? _pendingRedactionWorkspace;
    private long _pendingRedactionRevision;
    private long _redactionPlanGeneration;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ActiveEditorTool))]
    [NotifyPropertyChangedFor(nameof(EditorInstruction))]
    private PdfEditorToolChoice _selectedEditorToolChoice = AvailableEditorTools[0];

    [ObservableProperty]
    private string _editorText = "Review note";

    [ObservableProperty]
    private string _editorAuthor = Environment.UserName;

    [ObservableProperty]
    private string _editorColorHex = "#E5484D";

    [ObservableProperty]
    private string _editorStampName = "Approved";

    [ObservableProperty]
    private string _editorLinkUri = "https://";

    [ObservableProperty]
    private double _editorFontSize = 14D;

    [ObservableProperty]
    private string _redactionRemovedMarker = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPendingRedaction))]
    private string? _pendingRedactionSummary;

    [ObservableProperty]
    private string _watermarkText = "CONFIDENTIAL";

    [ObservableProperty]
    private PdfFormFieldViewModel? _selectedFormField;

    [ObservableProperty]
    private string _formFieldValue = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelectedAnnotation))]
    private int? _selectedAnnotationObjectNumber;

    [ObservableProperty]
    private string? _selectedAnnotationSummary;

    [ObservableProperty]
    private string _selectedAnnotationContents = string.Empty;

    [ObservableProperty]
    private string _selectedAnnotationAuthor = string.Empty;

    [ObservableProperty]
    private string _annotationReplyText = string.Empty;

    public ObservableCollection<PdfEditorToolChoice> EditorTools { get; } = new(AvailableEditorTools);

    public ObservableCollection<PdfFormFieldViewModel> FormFields { get; } = new();

    public PdfEditorTool ActiveEditorTool => SelectedEditorToolChoice.Tool;

    public string EditorInstruction => SelectedEditorToolChoice.Hint;

    public bool HasPendingRedaction => !string.IsNullOrWhiteSpace(PendingRedactionSummary);

    public bool HasFormFields => FormFields.Count > 0;

    public bool HasSelectedAnnotation => SelectedAnnotationObjectNumber.HasValue;

    public bool CanEditAnnotations => _workspace?.CanEditAnnotations == true;

    public bool CanEditPageContent => _workspace?.CanEditPageContent == true;

    public bool CanRedact => _workspace?.CanRedact == true;

    public bool CanFillForms => _workspace?.CanFillForms == true && SelectedFormField?.IsReadOnly == false;

    public bool CanFlattenForms => _workspace?.CanFlattenForms == true;

    public bool CanFillAndFlattenForms => CanFillForms && CanFlattenForms;

    partial void OnSelectedEditorToolChoiceChanged(PdfEditorToolChoice value) {
        foreach (PdfPageViewModel page in Pages) page.EditorTool = value.Tool;
        if (value.Tool != PdfEditorTool.Redact) CancelPendingRedaction();
    }

    partial void OnSelectedFormFieldChanged(PdfFormFieldViewModel? value) {
        FormFieldValue = value?.Value ?? string.Empty;
        OnPropertyChanged(nameof(CanFillForms));
        OnPropertyChanged(nameof(CanFillAndFlattenForms));
    }

    [RelayCommand]
    private void SelectEditorTool(string? label) {
        PdfEditorToolChoice? choice = EditorTools.FirstOrDefault(tool =>
            string.Equals(tool.Label, label, StringComparison.OrdinalIgnoreCase));
        if (choice is not null) SelectedEditorToolChoice = choice;
    }

    private async void OnPageEditorGestureCompleted(PdfEditorGesture gesture) {
        bool acceptsEditorGesture = DocumentMode is StudioDocumentMode.Annotate or StudioDocumentMode.Edit ||
                                    DocumentMode == StudioDocumentMode.Protect && ActiveEditorTool == PdfEditorTool.Redact;
        if (_workspace is null ||
            !acceptsEditorGesture ||
            ActiveEditorTool == PdfEditorTool.Select ||
            IsWorkspaceBusy) return;
        PdfWorkspace workspace = _workspace;
        long revision = workspace.Revision;
        PdfEditorTool tool = ActiveEditorTool;
        PdfEditorProperties properties = CreateEditorProperties();
        ErrorMessage = null;
        if (tool == PdfEditorTool.Redact) {
            if (!CanRedact) {
                ErrorMessage = "This document cannot be safely redacted under its current security and rewrite policy.";
                return;
            }
            CancelPendingRedaction();
            long generation = _redactionPlanGeneration;
            PdfRedactionPlan? plan = null;
            bool succeeded = await RunStandaloneAsync(async token => {
                OperationStatus = "Planning redaction";
                plan = await workspace.PlanRedactionAsync(gesture, properties, token).ConfigureAwait(true);
                token.ThrowIfCancellationRequested();
            }, CancellationToken.None).ConfigureAwait(true);
            if (!succeeded || plan is null) return;
            if (generation != _redactionPlanGeneration ||
                !ReferenceEquals(_workspace, workspace) ||
                workspace.Revision != revision ||
                ActiveEditorTool != PdfEditorTool.Redact) {
                OperationStatus = "The document changed before the redaction preview was ready. Draw the area again.";
                return;
            }
            int textMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.TextBlock);
            int imageMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.ImagePlacement);
            int annotationMatches = plan.Matches.Count(static match => match.Kind == PdfRedactionMatchKind.Annotation);
            _pendingRedaction = gesture;
            _pendingRedactionPlan = plan;
            _pendingRedactionWorkspace = workspace;
            _pendingRedactionRevision = revision;
            SetPendingRedactionArea(gesture);
            PendingRedactionSummary = $"Page {gesture.PageNumber}: {textMatches} text, {imageMatches} image, and {annotationMatches} annotation match(es). Intersecting images are removed as whole placements. Review the area, then apply permanent verified redaction.";
            return;
        }

        try {
            byte[]? imageBytes = null;
            if (tool == PdfEditorTool.AddImage) {
                string? path = await _pickImage(CancellationToken.None).ConfigureAwait(true);
                if (string.IsNullOrWhiteSpace(path)) return;
                imageBytes = await File.ReadAllBytesAsync(path).ConfigureAwait(true);
                if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) {
                    OperationStatus = "The document changed while the image was being selected. Place the image again.";
                    return;
                }
            }
            properties = properties with { ImageBytes = imageBytes };
            bool succeeded = await RunMutationAsync(
                token => workspace.ApplyEditorGestureAsync(tool, gesture, properties, token, CreateProgress()),
                CancellationToken.None).ConfigureAwait(true);
            if (succeeded) OperationStatus = "Edit added. Save when ready.";
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
        }
    }

    private void OnPageAnnotationSelected(PdfEditorSelection? selection) {
        if (_workspace is null || selection is null) {
            ClearAnnotationSelection();
            return;
        }
        PdfAnnotation? annotation = _workspace.DocumentInfo.Annotations.FirstOrDefault(candidate =>
            candidate.ObjectNumber == selection.ObjectNumber && candidate.PageNumber == selection.PageNumber);
        if (annotation is null) {
            ClearAnnotationSelection();
            return;
        }
        SelectedAnnotationObjectNumber = selection.ObjectNumber;
        SelectedAnnotationSummary = $"{selection.Subtype} · page {selection.PageNumber} · object {selection.ObjectNumber}";
        SelectedAnnotationContents = annotation.Contents ?? string.Empty;
        SelectedAnnotationAuthor = annotation.Title ?? string.Empty;
        foreach (PdfPageViewModel page in Pages) {
            page.SelectedAnnotationObjectNumber = page.PageNumber == selection.PageNumber ? selection.ObjectNumber : null;
        }
    }

    [RelayCommand]
    private async Task ApplyPendingRedactionAsync(CancellationToken cancellationToken) {
        if (_workspace is null ||
            _pendingRedaction is null ||
            _pendingRedactionPlan is null ||
            _pendingRedactionWorkspace is null) return;
        PdfWorkspace workspace = _pendingRedactionWorkspace;
        PdfRedactionPlan plan = _pendingRedactionPlan;
        long revision = _pendingRedactionRevision;
        if (!ReferenceEquals(_workspace, workspace) || workspace.Revision != revision) {
            CancelPendingRedaction();
            ErrorMessage = "The document changed after this redaction was reviewed. Draw and review the area again.";
            return;
        }
        PdfVerifiedRedactionResult? proof = null;
        bool succeeded = await RunMutationAsync(async token => {
            proof = await workspace.ApplyVerifiedRedactionAsync(
                plan,
                revision,
                RedactionRemovedMarker,
                token,
                CreateProgress()).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        if (!succeeded || proof is null) return;
        OperationStatus = proof.Verification.Summary + $" Removed {proof.Plan.Matches.Count} intersecting item(s).";
    }

    [RelayCommand]
    private void CancelPendingRedaction() {
        _redactionPlanGeneration++;
        _pendingRedaction = null;
        _pendingRedactionPlan = null;
        _pendingRedactionWorkspace = null;
        _pendingRedactionRevision = 0;
        PendingRedactionSummary = null;
        SetPendingRedactionArea(null);
    }

    private void SetPendingRedactionArea(PdfEditorGesture? gesture) {
        foreach (PdfPageViewModel page in Pages) {
            page.PendingRedactionArea = gesture is not null && page.PageNumber == gesture.PageNumber
                ? new Rect(gesture.Left, gesture.Top, gesture.Right - gesture.Left, gesture.Bottom - gesture.Top)
                : null;
        }
    }

    [RelayCommand]
    private async Task FillFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedFormField is null) return;
        await RunMutationAsync(
            token => _workspace.FillFormFieldAsync(SelectedFormField.Name, FormFieldValue, flatten: false, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FillAndFlattenFormFieldAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedFormField is null) return;
        await RunMutationAsync(
            token => _workspace.FillFormFieldAsync(SelectedFormField.Name, FormFieldValue, flatten: true, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task FlattenAllFormFieldsAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.FlattenFormFieldsAsync(token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ApplyWatermarkAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.ApplyWatermarkAsync(WatermarkText, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ApplyPageNumbersAsync(CancellationToken cancellationToken) {
        if (_workspace is null) return;
        await RunMutationAsync(
            token => _workspace.ApplyPageNumbersAsync(token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task UpdateSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        PdfColor color = ParseColor(EditorColorHex);
        string contents = SelectedAnnotationContents;
        string author = SelectedAnnotationAuthor;
        ClearAnnotationSelection();
        await RunMutationAsync(
            token => _workspace.UpdateAnnotationAsync(objectNumber, contents, author, color, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task ReplyToSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        string reply = AnnotationReplyText;
        PdfColor color = ParseColor(EditorColorHex);
        ClearAnnotationSelection();
        await RunMutationAsync(
            token => _workspace.AddAnnotationReplyAsync(objectNumber, reply, EditorAuthor, color, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        AnnotationReplyText = string.Empty;
    }

    [RelayCommand]
    private async Task FlattenSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        ClearAnnotationSelection();
        await RunMutationAsync(
            token => _workspace.FlattenAnnotationAsync(objectNumber, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task RemoveSelectedAnnotationAsync(CancellationToken cancellationToken) {
        if (_workspace is null || SelectedAnnotationObjectNumber is not int objectNumber) return;
        ClearAnnotationSelection();
        await RunMutationAsync(
            token => _workspace.RemoveAnnotationAsync(objectNumber, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
    }

    private PdfEditorProperties CreateEditorProperties() => new(
        EditorText ?? string.Empty,
        EditorAuthor ?? string.Empty,
        ParseColor(EditorColorHex),
        string.IsNullOrWhiteSpace(EditorStampName) ? "Approved" : EditorStampName.Trim(),
        EditorLinkUri ?? string.Empty,
        Math.Clamp(EditorFontSize, 4D, 144D));

    private void RebuildFormFields() {
        string? selectedName = SelectedFormField?.Name;
        FormFields.Clear();
        if (_workspace is not null) {
            foreach (PdfFormField field in _workspace.DocumentInfo.FormFields.Where(static field => !string.IsNullOrWhiteSpace(field.Name))) {
                FormFields.Add(new PdfFormFieldViewModel(
                    field.Name!,
                    field.Kind.ToString(),
                    field.Value ?? string.Empty,
                    field.IsReadOnly,
                    field.PageNumbers));
            }
        }
        SelectedFormField = FormFields.FirstOrDefault(field => string.Equals(field.Name, selectedName, StringComparison.Ordinal))
            ?? FormFields.FirstOrDefault();
        OnPropertyChanged(nameof(HasFormFields));
        OnPropertyChanged(nameof(CanFillForms));
        OnPropertyChanged(nameof(CanFlattenForms));
        OnPropertyChanged(nameof(CanFillAndFlattenForms));
    }

    private void ClearAnnotationSelection() {
        SelectedAnnotationObjectNumber = null;
        SelectedAnnotationSummary = null;
        SelectedAnnotationContents = string.Empty;
        SelectedAnnotationAuthor = string.Empty;
        foreach (PdfPageViewModel page in Pages) page.SelectedAnnotationObjectNumber = null;
    }

    private static PdfColor ParseColor(string? value) {
        string hex = (value ?? string.Empty).Trim().TrimStart('#');
        if (hex.Length != 6 ||
            !byte.TryParse(hex.AsSpan(0, 2), System.Globalization.NumberStyles.HexNumber, null, out byte red) ||
            !byte.TryParse(hex.AsSpan(2, 2), System.Globalization.NumberStyles.HexNumber, null, out byte green) ||
            !byte.TryParse(hex.AsSpan(4, 2), System.Globalization.NumberStyles.HexNumber, null, out byte blue)) {
            throw new FormatException("Editor color must be a six-digit hex value such as #E5484D.");
        }
        return PdfColor.FromRgb(red, green, blue);
    }
}
